# Glaronia Informatik AG
# Rikardo Stoilov
# 22.08.2024

import re
import sys
import pdfplumber
from pathlib import Path
from datetime import datetime

REPORTING_INTERVAL_PAT = re.compile(r"Reporting Interval:\s*.+?\.(\d{2}\.\d{4})\s*[–—-]\s*.+?\.\d{2}\.\d{4}")
COMPANY_PAT = re.compile(r"^(.*) \(Support ID: \d{8}\)$")
VM_PAT = re.compile(r"^VM\s+(\d+)\s+(\d+)")
WORKSTATION_PAT = re.compile(r"^Workstation\s+(\d+)\s+(\d+)")
SERVER_PAT = re.compile(r"^Server\s+(\d+)\s+(\d+)")
FILE_SHARE_PAT = re.compile(r"^File Share\s+(\d+)\s+(\d+)")
CLOUD_VM_PAT = re.compile(r"^Cloud VM\s+(\d+)\s+(\d+)")
CLOUD_BACKUP_VM_PAT = re.compile(r"^Cloud Backup \(VM\)\s+(\d+)\s+(\d+)")

def extract_rows_from_pdf(pdf_file):
    rows: list[dict] = []
    date = None

    pdf_path = Path(pdf_file)
    try:
        rows, date = parse_single_pdf(pdf_path)
        print(f"✓ Verarbeitet: {pdf_path.name}  -> {len(rows)} Einträge")
    except Exception as e:
        print(f"⚠️ Fehler bei {pdf_path.name}: {e}")
        sys.exit(1)

    return rows, date

def parse_single_pdf(pdf_path: Path):
    date = None
    rows = []
    current_company = None
    patterns = {
        "VM": VM_PAT,
        "Workstation": WORKSTATION_PAT,
        "Server": SERVER_PAT,
        "File Share": FILE_SHARE_PAT,
        "Cloud VM": CLOUD_VM_PAT,
        "Cloud Backup VM": CLOUD_BACKUP_VM_PAT
    }
    matches = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [l.strip() for l in text.split("\n") if l.strip()]

            for line in lines:
                date_match = REPORTING_INTERVAL_PAT.search(line)
                if date_match:
                    date = datetime.strptime(date_match.group(1), "%m.%Y")
                    date = date.strftime("%Y-%m")

                company_match = COMPANY_PAT.search(line)
                if company_match:
                    if current_company:
                        rows = append_row(rows, pdf_path, current_company, matches)
                    current_company = company_match.group(1).strip()
                    matches = {}
                    continue

                if current_company:
                    for key, pat in patterns.items():
                        match = pat.search(line)
                        if match:
                            matches[key] = int(match.group(1)) + int(match.group(2))
                            break
                    continue

        # Aller Letzten Block speichern
        if current_company:
            rows = append_row(rows, pdf_path, current_company, matches)

    return rows, date

def append_row(rows, pdf_path, current_company, matches):
    rows.append({
        "Quelle": pdf_path.name,
        "Firma": current_company,
        "VM": matches.get("VM", None),
        "Workstation": matches.get("Workstation", None),
        "Server": matches.get("Server", None),
        "File Share": matches.get("File Share", None),
        "Cloud VM": matches.get("Cloud VM", None),
        "Cloud Backup VM": matches.get("Cloud Backup VM", None)
    })

    return rows
