import json
import os
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo

ROOT = Path(__file__).resolve().parent
DATA_FILE = ROOT / "jobs.json"
XLSX_FILE = ROOT / "MuleSoft_Job_Tracker.xlsx"
IST = timezone(timedelta(hours=5, minutes=30))

HEADERS = [
    "Date & Time Found",
    "Company Name",
    "Job Title",
    "Location / Remote Status",
    "Match Score / Relevance Note",
    "Direct Application Link",
]

PROFILE_TERMS = {
    "mulesoft": 20, "mule 4": 12, "anypoint": 10, "dataweave": 10,
    "raml": 7, "apikit": 6, "cloudhub": 6, "api manager": 5,
    "runtime manager": 4, "munit": 5, "salesforce": 3, "sftp": 2,
    "anypoint mq": 2, "oauth": 2, "jwt": 2, "postman": 2, "soap": 2,
    "sql": 2, "mysql": 2, "object store": 2
}
POSITIVE_ROLE_WORDS = ("mulesoft", "mule soft", "integration", "api developer", "api engineer")
SENIOR_WORDS = ("architect", "principal", "staff", "director", "head of", "lead architect")
SENIOR_6_PLUS = re.compile(r"(?<!\d)(?:6|7|8|9|10|1[1-9]|20)\s*\+?\s*(?:years|yrs)", re.I)
EXPERIENCE_RANGE = re.compile(r"(?<!\d)(\d+)\s*[-–to]+\s*(\d+)\s*(?:years|yrs)", re.I)
YEARS_REQ = re.compile(r"(?<!\d)(\d+)\s*\+?\s*(?:years|yrs)", re.I)

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (compatible; MuleSoftJobTracker/1.0; +https://github.com/sonii-shivansh/mulesoft-tracker)"
})

def now_ist():
    return datetime.now(IST).strftime("%Y-%m-%d %H:%M IST")

def normalize_url(url):
    if not url:
        return ""
    return url.split("#")[0].strip().rstrip("/")

def job_key(job):
    url = normalize_url(job.get("url", ""))
    if url:
        return url.lower()
    return "|".join([
        job.get("company", "").strip().lower(),
        job.get("title", "").strip().lower(),
        job.get("location", "").strip().lower(),
    ])

def score_job(title, description):
    text = f"{title} {description}".lower()
    if not any(x in text for x in POSITIVE_ROLE_WORDS):
        return None
    if any(x in title.lower() for x in SENIOR_WORDS):
        return None
    if SENIOR_6_PLUS.search(text):
        return None

    score = 35
    for term, points in PROFILE_TERMS.items():
        if term in text:
            score += points

    years = [int(x) for x in YEARS_REQ.findall(text)]
    if years:
        minimum = min(years)
        if minimum >= 6:
            return None
        if minimum <= 4:
            score += 8

    return min(score, 99)

def relevance_note(title, description, score):
    text = description.lower()
    found = [term for term in PROFILE_TERMS if term in text]
    top = ", ".join(found[:9])
    return f"{score}% — Strong profile overlap: {top}." if top else f"{score}% — MuleSoft/integration role matching target."

def fetch_greenhouse_jobs(board_token, company_name):
    url = f"https://boards-api.greenhouse.io/v1/boards/{board_token}/jobs?content=true"
    r = session.get(url, timeout=25)
    r.raise_for_status()
    out = []
    for item in r.json().get("jobs", []):
        title = item.get("title", "")
        desc = BeautifulSoup(item.get("content", ""), "html.parser").get_text(" ", strip=True)
        score = score_job(title, desc)
        if score is None:
            continue
        out.append({
            "date_time_found": now_ist(),
            "company": company_name,
            "title": title,
            "location": (item.get("location") or {}).get("name", "Not specified"),
            "match": relevance_note(title, desc, score),
            "url": item.get("absolute_url", ""),
            "source": "Greenhouse"
        })
    return out

def fetch_lever_jobs(company_slug, company_name):
    url = f"https://api.lever.co/v0/postings/{company_slug}?mode=json"
    r = session.get(url, timeout=25)
    r.raise_for_status()
    out = []
    for item in r.json():
        title = item.get("text", "")
        desc = BeautifulSoup(
            (item.get("descriptionPlain") or "") + " " +
            (item.get("additionalPlain") or ""),
            "html.parser"
        ).get_text(" ", strip=True)
        score = score_job(title, desc)
        if score is None:
            continue
        categories = item.get("categories") or {}
        location = categories.get("location") or "Not specified"
        out.append({
            "date_time_found": now_ist(),
            "company": company_name,
            "title": title,
            "location": location,
            "match": relevance_note(title, desc, score),
            "url": item.get("hostedUrl", ""),
            "source": "Lever"
        })
    return out

def fetch_rss(url, company_hint=None):
    r = session.get(url, timeout=25)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "xml")
    out = []
    for item in soup.find_all(["item", "entry"]):
        title = (item.find("title").get_text(" ", strip=True) if item.find("title") else "")
        desc_node = item.find("description") or item.find("summary") or item.find("content")
        desc = desc_node.get_text(" ", strip=True) if desc_node else ""
        link_node = item.find("link")
        link = link_node.get("href") if link_node and link_node.get("href") else (link_node.get_text(strip=True) if link_node else "")
        score = score_job(title, desc)
        if score is None:
            continue
        company = company_hint or "Unknown"
        out.append({
            "date_time_found": now_ist(),
            "company": company,
            "title": title,
            "location": "See posting",
            "match": relevance_note(title, desc, score),
            "url": link,
            "source": "RSS"
        })
    return out

def discover():
    jobs = []
    # Optional direct ATS feeds. Add comma-separated board tokens/slugs through GitHub secrets.
    greenhouse = os.getenv("GREENHOUSE_BOARDS", "")
    for entry in filter(None, greenhouse.split(",")):
        try:
            token, company = [x.strip() for x in entry.split("|", 1)]
            jobs.extend(fetch_greenhouse_jobs(token, company))
        except Exception as exc:
            print(f"Greenhouse provider failed for {entry}: {exc}")

    lever = os.getenv("LEVER_BOARDS", "")
    for entry in filter(None, lever.split(",")):
        try:
            slug, company = [x.strip() for x in entry.split("|", 1)]
            jobs.extend(fetch_lever_jobs(slug, company))
        except Exception as exc:
            print(f"Lever provider failed for {entry}: {exc}")

    rss = os.getenv("JOB_RSS_URLS", "")
    for entry in filter(None, rss.split(",")):
        try:
            parts = entry.split("|", 1)
            jobs.extend(fetch_rss(parts[0].strip(), parts[1].strip() if len(parts) > 1 else None))
        except Exception as exc:
            print(f"RSS provider failed for {entry}: {exc}")

    return jobs

def load_jobs():
    if not DATA_FILE.exists():
        return []
    return json.loads(DATA_FILE.read_text(encoding="utf-8"))

def save_jobs(jobs):
    DATA_FILE.write_text(json.dumps(jobs, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

def append_unique(existing, discovered):
    keys = {job_key(j) for j in existing}
    added = []
    for job in discovered:
        key = job_key(job)
        if not key or key in keys:
            continue
        keys.add(key)
        existing.append(job)
        added.append(job)
    return added

def build_xlsx(jobs):
    wb = Workbook()
    ws = wb.active
    ws.title = "MuleSoft Jobs"
    ws.append(HEADERS)

    for job in jobs:
        ws.append([
            job.get("date_time_found", ""),
            job.get("company", ""),
            job.get("title", ""),
            job.get("location", ""),
            job.get("match", ""),
            job.get("url", ""),
        ])

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E1F2")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=thin)

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = Border(bottom=thin)
        if row[5].value:
            row[5].hyperlink = row[5].value
            row[5].style = "Hyperlink"

    widths = [21, 30, 45, 45, 90, 75]
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[chr(64 + idx)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 32

    if ws.max_row >= 2:
        table = Table(displayName="MuleSoftJobTracker", ref=f"A1:F{ws.max_row}")
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        ws.add_table(table)

    info = wb.create_sheet("Tracker Info")
    info.append(["Field", "Value"])
    info_rows = [
        ("Profile", "MuleSoft Developer / Integration Engineer — approximately 2–4 years"),
        ("Core stack", "Mule 4, Anypoint Platform, RAML, APIKit, DataWeave 2.0, API-led connectivity"),
        ("Filter", "Exclude roles explicitly requiring 6+ years or architect/principal/director-level roles"),
        ("Deduplication", "Application URL first; otherwise company + title + location"),
        ("Automation", "GitHub Actions scheduled hourly"),
        ("Providers", "Direct ATS APIs (Greenhouse/Lever) and configured RSS feeds; provider configuration is intentionally secret/config driven"),
        ("Last generated", now_ist()),
    ]
    for row in info_rows:
        info.append(row)
    info.column_dimensions["A"].width = 28
    info.column_dimensions["B"].width = 110
    for cell in info[1]:
        cell.fill = header_fill
        cell.font = header_font
    for row in info.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    wb.save(XLSX_FILE)

def main():
    existing = load_jobs()
    discovered = discover()
    added = append_unique(existing, discovered)
    save_jobs(existing)
    build_xlsx(existing)

    print(f"TOTAL_JOBS={len(existing)}")
    print(f"NEW_JOBS={len(added)}")
    for job in added:
        print("NEW_JOB|" + json.dumps(job, ensure_ascii=False))

if __name__ == "__main__":
    main()
