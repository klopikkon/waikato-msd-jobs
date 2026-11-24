"""Fetch active jobs from MyMSD (new.my.msd.govt.nz) and export them to Excel + JSON.

Usage:
    python fetch_msd_jobs.py
    python fetch_msd_jobs.py --docs
    python fetch_msd_jobs.py --url "https://new.my.msd.govt.nz/en/jobs/search?l=..." --output waikato_jobs.xlsx
    python fetch_msd_jobs.py --region "Waikato Region" --output waikato_jobs.xlsx
    python fetch_msd_jobs.py --district "Hamilton City" --output hamilton_jobs.xlsx
"""

from __future__ import annotations

import argparse
import base64
import html
import json
import re
import urllib.parse
import urllib.request
import uuid
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone
from html.parser import HTMLParser
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

API_URL = "https://new.my.msd.govt.nz/bin/sep/vacancies/search"
DETAILS_API_URL = "https://new.my.msd.govt.nz/bin/sep/vacancies/details"
DEFAULT_SEARCH_URL = (
    "https://new.my.msd.govt.nz/en/jobs/search"
    "?l=W3siZGVzYyI6WyJXYWlrYXRvJTIwUmVnaW9uIiwiTm9ydGglMjBJc2xhbmQiXSwicmVnaW9uIjp0cnVlfV0%3D"
)
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
)


class _HTMLToText(HTMLParser):
    """Convert MyMSD HTML vacancy descriptions into clean, Excel-friendly plain text."""

    def __init__(self) -> None:
        super().__init__()
        self._parts: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "li":
            self._parts.append("\n• ")
        elif tag == "br":
            self._parts.append("\n")
        elif tag in ("p", "div", "h1", "h2", "h3", "h4"):
            if self._parts and not self._parts[-1].endswith("\n"):
                self._parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in ("p", "div", "ul", "ol", "h1", "h2", "h3", "h4"):
            self._parts.append("\n")

    def handle_data(self, data: str) -> None:
        self._parts.append(data)

    def get_text(self) -> str:
        raw = html.unescape("".join(self._parts))
        # Collapse excessive blank lines and trailing spaces per line
        lines = [line.strip() for line in raw.splitlines()]
        cleaned = re.sub(r"\n{3,}", "\n\n", "\n".join(lines)).strip()
        # Excel cells cap at 32,767 characters
        return cleaned[:32000]


def html_to_plain_text(raw_html: str) -> str:
    if not raw_html:
        return ""
    parser = _HTMLToText()
    parser.feed(raw_html)
    parser.close()
    return parser.get_text()


def payload_from_search_url(url: str) -> dict[str, Any]:
    """Decode a MyMSD /en/jobs/search URL into the /bin/sep/vacancies/search JSON body."""
    parsed = urllib.parse.urlparse(url)
    params = urllib.parse.parse_qs(parsed.query)

    payload: dict[str, Any] = {}
    filters: dict[str, Any] = {}

    if "q" in params and params["q"][0].strip():
        payload["searchTerms"] = params["q"][0].strip()

    if "c" in params and params["c"]:
        filters["occupationGroups"] = [c for c in params["c"] if c.strip()]

    if "l" in params and params["l"][0]:
        raw_l = urllib.parse.unquote(params["l"][0])
        padding = "=" * (-len(raw_l) % 4)
        decoded = json.loads(base64.b64decode(raw_l + padding).decode("utf-8"))
        regions: list[str] = []
        districts: list[str] = []
        for item in decoded:
            desc = item.get("desc") or []
            if not desc:
                continue
            name = urllib.parse.unquote(desc[0])
            if item.get("region"):
                regions.append(name)
            else:
                districts.append(name)
        locations: dict[str, list[str]] = {}
        if regions:
            locations["regions"] = regions
        if districts:
            locations["districts"] = districts
        if locations:
            filters["locations"] = locations

    if filters:
        payload["filters"] = filters

    if not payload:
        raise ValueError(
            "No search terms, category, or location filters found in the URL. "
            "MyMSD requires at least one filter."
        )
    return payload


def fetch_vacancies(payload: dict[str, Any]) -> list[dict[str, Any]]:
    """POST to the MyMSD vacancies search API and return all matching vacancies."""
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        API_URL,
        data=data,
        method="POST",
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json",
            "Origin": "https://new.my.msd.govt.nz",
            "Referer": "https://new.my.msd.govt.nz/en/jobs/search",
            "x-transaction-id": str(uuid.uuid4()),
            "User-Agent": USER_AGENT,
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = json.loads(resp.read().decode("utf-8"))

    results = body.get("results") or {}
    return results.get("suitableVacancies") or []


def fetch_vacancy_description(vacancy_id: str) -> str:
    """GET individual vacancy details and return the plain-text position description."""
    if not vacancy_id:
        return ""
    url = f"{DETAILS_API_URL}?vacancyId={urllib.parse.quote(vacancy_id)}"
    req = urllib.request.Request(
        url,
        method="GET",
        headers={
            "Accept": "application/json",
            "Origin": "https://new.my.msd.govt.nz",
            "Referer": f"https://new.my.msd.govt.nz/en/jobs/vacancy/{vacancy_id}",
            "x-transaction-id": str(uuid.uuid4()),
            "User-Agent": USER_AGENT,
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            body = json.loads(resp.read().decode("utf-8"))
        details = (body.get("results") or {}).get("vacancyDetails") or {}
        return html_to_plain_text(details.get("vacancyDescription") or "")
    except Exception:
        return ""


def fetch_all_descriptions(vacancy_ids: list[str], max_workers: int = 5) -> dict[str, str]:
    """Fetch position descriptions for all vacancy IDs with polite concurrency."""
    descriptions: dict[str, str] = {}
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        futures = {pool.submit(fetch_vacancy_description, vid): vid for vid in vacancy_ids if vid}
        for future in as_completed(futures):
            vid = futures[future]
            descriptions[vid] = future.result()
    return descriptions


def format_pay(job: dict[str, Any]) -> str:
    if not job.get("displayPayOnListing"):
        return "To be discussed"
    min_pay = job.get("minimumPay")
    max_pay = job.get("maximumPay")
    pay_type = job.get("payType") or ""
    if min_pay is not None and max_pay is not None and min_pay != max_pay:
        return f"${min_pay:,.2f} - ${max_pay:,.2f} ({pay_type})".strip()
    if min_pay is not None:
        return f"${min_pay:,.2f} ({pay_type})".strip()
    if max_pay is not None:
        return f"${max_pay:,.2f} ({pay_type})".strip()
    return "To be discussed"


def normalize_job(job: dict[str, Any]) -> dict[str, Any]:
    locs = job.get("locations") or {}
    posting = job.get("vacancyPosting") or {}
    occ = job.get("occupation") or {}
    occ_group = occ.get("occupationGroup") or {}
    industry = job.get("industry") or {}

    vacancy_id = job.get("vacancyId", "")
    posting_url = posting.get("postingURL") or (
        f"https://new.my.msd.govt.nz/en/jobs/vacancy/{vacancy_id}" if vacancy_id else ""
    )

    return {
        "vacancyId": vacancy_id,
        "title": job.get("vacancyTitle", ""),
        "category": occ_group.get("name", ""),
        "occupation": occ.get("name", ""),
        "industry": industry.get("value", ""),
        "contractType": job.get("vacancyType", ""),
        "hours": job.get("fullPartTime", ""),
        "workplace": job.get("workPlaceArrangement", ""),
        "pay": format_pay(job),
        "minPay": job.get("minimumPay") if job.get("displayPayOnListing") else None,
        "maxPay": job.get("maximumPay") if job.get("displayPayOnListing") else None,
        "payType": job.get("payType", "") if job.get("displayPayOnListing") else "",
        "regions": ", ".join(locs.get("regions") or []),
        "districts": ", ".join(locs.get("districts") or []),
        "suburbs": ", ".join(locs.get("suburbs") or []),
        "startDate": posting.get("startDate", ""),
        "endDate": posting.get("endDate", ""),
        "status": job.get("vacancyStatus", ""),
        "url": posting_url,
    }


def export_to_excel(
    jobs: list[dict[str, Any]],
    descriptions: dict[str, str],
    output_path: Path,
) -> None:
    """Write normalized jobs (including Job ID and Position Description) to Excel."""
    wb = Workbook()
    ws = wb.active
    ws.title = "MyMSD Jobs"

    headers = [
        "Job ID",
        "Title",
        "Position Description",
        "Category",
        "Occupation",
        "Industry",
        "Contract Type",
        "Hours",
        "Workplace",
        "Pay",
        "Min Pay",
        "Max Pay",
        "Pay Type",
        "Regions",
        "Districts",
        "Suburbs",
        "Start Date",
        "End Date",
        "Status",
        "URL",
    ]

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="left", vertical="center")

    link_font = Font(color="0563C1", underline="single")
    top_align = Alignment(vertical="top")
    wrap_align = Alignment(vertical="top", wrap_text=True)

    for row_idx, job in enumerate(jobs, start=2):
        desc = descriptions.get(job["vacancyId"], "")
        row = [
            job["vacancyId"],
            job["title"],
            desc,
            job["category"],
            job["occupation"],
            job["industry"],
            job["contractType"],
            job["hours"],
            job["workplace"],
            job["pay"],
            job["minPay"],
            job["maxPay"],
            job["payType"],
            job["regions"],
            job["districts"],
            job["suburbs"],
            job["startDate"],
            job["endDate"],
            job["status"],
            job["url"],
        ]
        ws.append(row)

        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.alignment = wrap_align if col_idx == 3 else top_align

        if job["url"]:
            url_cell = ws.cell(row=row_idx, column=len(headers))
            url_cell.hyperlink = job["url"]
            url_cell.font = link_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(jobs) + 1}"

    for col_idx in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        if col_idx == 3:  # Position Description column
            ws.column_dimensions[col_letter].width = 65
            continue
        max_len = max(
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, len(jobs) + 2)
        )
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def now_nz() -> datetime:
    """Return current New Zealand time (falls back to UTC+13 if tzdata is missing)."""
    try:
        return datetime.now(ZoneInfo("Pacific/Auckland"))
    except ZoneInfoNotFoundError:
        return datetime.now(timezone(timedelta(hours=13)))


def export_to_json(jobs: list[dict[str, Any]], output_path: Path, excel_filename: str = "") -> None:
    """Write normalized jobs (without bulky descriptions) and timestamp to JSON for the web UI."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload: dict[str, Any] = {
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        "total": len(jobs),
        "jobs": jobs,
    }
    if excel_filename:
        payload["excelFile"] = excel_filename
    output_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Fetch active jobs from MyMSD and export them to Excel."
    )
    parser.add_argument(
        "--url",
        default=DEFAULT_SEARCH_URL,
        help="MyMSD /en/jobs/search URL (defaults to Waikato Region).",
    )
    parser.add_argument(
        "--region",
        action="append",
        help='Override with one or more region names (e.g. --region "Waikato Region").',
    )
    parser.add_argument(
        "--district",
        action="append",
        help='Override with one or more district names (e.g. --district "Hamilton City").',
    )
    parser.add_argument(
        "--output",
        default="waikato_jobs.xlsx",
        help="Output Excel file path (default: waikato_jobs.xlsx).",
    )
    parser.add_argument(
        "--json-output",
        default="",
        help="Optional JSON output path for the GitHub Pages web viewer.",
    )
    parser.add_argument(
        "--docs",
        action="store_true",
        help="Write docs/waikato_jobs.xlsx and docs/jobs.json for GitHub Pages.",
    )
    args = parser.parse_args()

    if args.region or args.district:
        locations: dict[str, list[str]] = {}
        if args.region:
            locations["regions"] = args.region
        if args.district:
            locations["districts"] = args.district
        payload: dict[str, Any] = {"filters": {"locations": locations}}
    else:
        payload = payload_from_search_url(args.url)

    print(f"Querying {API_URL} with payload: {json.dumps(payload)}")
    raw_vacancies = fetch_vacancies(payload)
    jobs = [normalize_job(j) for j in raw_vacancies]
    print(f"Fetched {len(jobs)} active vacancies.")

    print("Fetching position descriptions for Excel export...")
    descriptions = fetch_all_descriptions([j["vacancyId"] for j in jobs])
    fetched_count = sum(1 for v in descriptions.values() if v)
    print(f"Fetched {fetched_count}/{len(jobs)} position descriptions.")

    if args.docs:
        docs_dir = Path("docs").resolve()
        docs_dir.mkdir(parents=True, exist_ok=True)
        for old_xlsx in docs_dir.glob("waikato_jobs*.xlsx"):
            old_xlsx.unlink()

        stamp = now_nz().strftime("%Y-%m-%d_%H-%M_NZT")
        excel_filename = f"waikato_jobs_{stamp}.xlsx"
        excel_path = docs_dir / excel_filename
        json_path = docs_dir / "jobs.json"
        export_to_excel(jobs, descriptions, excel_path)
        export_to_json(jobs, json_path, excel_filename=excel_filename)
        print(f"Saved Excel file to: {excel_path}")
        print(f"Saved JSON file to: {json_path}")
    else:
        excel_path = Path(args.output).resolve()
        export_to_excel(jobs, descriptions, excel_path)
        print(f"Saved Excel file to: {excel_path}")
        if args.json_output:
            json_path = Path(args.json_output).resolve()
            export_to_json(jobs, json_path)
            print(f"Saved JSON file to: {json_path}")


if __name__ == "__main__":
    main()
