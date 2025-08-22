# Glaronia Informatik AG
# Rikardo Stoilov
# 22.08.2024

import re
import sys
import pdfplumber
from pathlib import Path
from datetime import datetime

REPORTING_INTERVAL_PAT = re.compile(r".+?\.(\d{2}\.\d{4})\s*[–—-]\s*.+?\.\d{2}\.\d{4}")
COMPANY_PAT = re.compile(r"Organization:\s*(.+)", re.IGNORECASE)
LICENSES_PAT = re.compile(r"Total number of licenses:\s*(\d+)", re.IGNORECASE)
NEW_LICENSES_PAT = re.compile(r"Total number of new licenses:\s*(\d+)", re.IGNORECASE)

def extract_rows_from_pdfs(pdf_files):
    all_rows: list[dict] = []
    all_dates: list = []

    for pdf_file in pdf_files:
        pdf_path = Path(pdf_file)
        try:
            rows, date = parse_single_pdf(pdf_path)
            all_rows.extend(rows)
            all_dates.append(date)
            print(f"✓ Verarbeitet: {pdf_path.name}  -> {len(rows)} Einträge")
        except Exception as e:
            print(f"⚠️ Fehler bei {pdf_path.name}: {e}")
            sys.exit(1)

    if not len(set(all_dates)) == 1:
        print(f"⚠️ Fehler, die ausgewählten PDFs sind nicht vom gleichen Zeitraum: {all_dates}")
        sys.exit(1)

    return all_rows, all_dates[0]

def parse_single_pdf(pdf_path: Path):
    rows = []
    current_company = None
    licenses = None
    new_licenses = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [l.strip() for l in text.split("\n") if l.strip()]

            for line in lines:
                date_match = REPORTING_INTERVAL_PAT.search(line)
                if date_match:
                    date = datetime.strptime(date_match.group(1), "%m.%Y")
                    date = date.strftime("%Y-%m")
                    continue

                company_match = COMPANY_PAT.search(line)
                if company_match:
                    if current_company:
                        rows = append_row(rows, pdf_path, current_company, licenses, new_licenses)
                    current_company = company_match.group(1).strip()
                    licenses = None
                    new_licenses = None
                    continue

                licenses_match = LICENSES_PAT.search(line)
                if licenses_match:
                    licenses = int(licenses_match.group(1))
                    continue

                new_licenses_match = NEW_LICENSES_PAT.search(line)
                if new_licenses_match:
                    new_licenses = int(new_licenses_match.group(1))
                    continue

        if current_company:
            rows = append_row(rows, pdf_path, current_company, licenses, new_licenses)

    return rows, date

def append_row(rows, pdf_path, current_company, licenses, new_licenses):
    rows.append({
        "Quelle": pdf_path.name,
        "Firma": current_company,
        "Lizenzen": licenses,
        "Neue_Lizenzen": new_licenses or 0,
        "Total_Lizenzen": licenses + (new_licenses or 0),
    })

    return rows
