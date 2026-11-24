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
import json
import urllib.parse
import urllib.request
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

API_URL = "https://new.my.msd.govt.nz/bin/sep/vacancies/search"
DEFAULT_SEARCH_URL = (
    "https://new.my.msd.govt.nz/en/jobs/search"
    "?l=W3siZGVzYyI6WyJXYWlrYXRvJTIwUmVnaW9uIiwiTm9ydGglMjBJc2xhbmQiXSwicmVnaW9uIjp0cnVlfV0%3D"
)


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
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        body = json.loads(resp.read().decode("utf-8"))

    results = body.get("results") or {}
    return results.get("suitableVacancies") or []


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


def export_to_excel(jobs: list[dict[str, Any]], output_path: Path) -> None:
    """Write normalized jobs to a formatted Excel workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = "MyMSD Jobs"

    headers = [
        "Vacancy ID",
        "Title",
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

    for row_idx, job in enumerate(jobs, start=2):
        row = [
            job["vacancyId"],
            job["title"],
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

        if job["url"]:
            url_cell = ws.cell(row=row_idx, column=len(headers))
            url_cell.hyperlink = job["url"]
            url_cell.font = link_font

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(jobs) + 1}"

    for col_idx in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(
            len(str(ws.cell(row=r, column=col_idx).value or ""))
            for r in range(1, len(jobs) + 2)
        )
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def export_to_json(jobs: list[dict[str, Any]], output_path: Path) -> None:
    """Write normalized jobs and update timestamp to JSON for the web UI."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "updatedAt": datetime.now(timezone.utc).isoformat(),
        "total": len(jobs),
        "jobs": jobs,
    }
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

    if args.docs:
        excel_path = Path("docs/waikato_jobs.xlsx").resolve()
        json_path = Path("docs/jobs.json").resolve()
        export_to_excel(jobs, excel_path)
        export_to_json(jobs, json_path)
        print(f"Saved Excel file to: {excel_path}")
        print(f"Saved JSON file to: {json_path}")
    else:
        excel_path = Path(args.output).resolve()
        export_to_excel(jobs, excel_path)
        print(f"Saved Excel file to: {excel_path}")
        if args.json_output:
            json_path = Path(args.json_output).resolve()
            export_to_json(jobs, json_path)
            print(f"Saved JSON file to: {json_path}")


if __name__ == "__main__":
    main()
