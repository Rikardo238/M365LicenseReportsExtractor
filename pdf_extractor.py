# Glaronia Informatik AG
# Rikardo Stoilov
# 20.08.2024

import sys
import re
import pdfplumber
import pandas as pd
from datetime import datetime
from pathlib import Path

# --- Muster (ignoriert Gross-/Kleinschreibung) ---
ORG_PAT = re.compile(r"Organization:\s*(.+)", re.IGNORECASE)
TOTAL_PAT = re.compile(r"Total number of licenses:\s*(\d+)", re.IGNORECASE)
NEW_PAT = re.compile(r"Total number of new licenses:\s*(\d+)", re.IGNORECASE)
REPORTING_INTERVAL_PAT = re.compile(r"(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})")

def get_reporting_interval(pdf_path):
    """Liest den Reporting Interval (Start- und Enddatum) von der ersten Seite."""
    pdf_first_page_text = None
    with pdfplumber.open(pdf_path) as pdf:
        if not pdf.pages:
            print(f"⚠️ Fehler: Das PDF {pdf_path.name} kann nicht geöffnet werden!")
            sys.exit(1)
        pdf_first_page_text = pdf.pages[0].extract_text() or ""

    match = REPORTING_INTERVAL_PAT.search(pdf_first_page_text)
    if not match:
        print(f"⚠️ Fehler: Ein Zeitraum konnte nicht wurden werden!")
        sys.exit(1)
    date = match.group(1)
    date = datetime.strptime(date, "%d.%m.%Y")
    date = date.strftime("%Y-%m")
    #all_months = ["Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sept", "Okt", "Nov", "Dez"]
    #return all_months[dt.month - 1]
    return date


def parse_single_pdf(pdf_path: Path) -> list[dict]:
    """Liest EIN PDF, überspringt die erste Seite und gibt eine Liste mit dicts je Organisation zurück."""
    rows = []
    current_org = None
    licenses = None
    new_licenses = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = [l.strip() for l in text.split("\n") if l.strip()]

            for line in lines:
                # Neue Organisation beginnt
                org_match = ORG_PAT.search(line)
                if org_match:
                    # Vorherige Organisation abschliessen (falls vorhanden)
                    if current_org:
                        rows.append({
                            "Quelle": pdf_path.name,
                            "Firma": current_org,
                            "Lizenzen": licenses,
                            "Neue_Lizenzen": new_licenses or 0,
                            "Total_Lizenzen": licenses + (new_licenses or 0),
                        })
                    current_org = org_match.group(1).strip()
                    licenses = None
                    new_licenses = None
                    continue

                # Gesamt-Lizenzen
                m_total = TOTAL_PAT.search(line)
                if m_total:
                    licenses = int(m_total.group(1))
                    continue

                # Neue Lizenzen
                m_new = NEW_PAT.search(line)
                if m_new:
                    new_licenses = int(m_new.group(1))
                    continue

        # Letzten Block speichern
        if current_org:
            rows.append({
                "Quelle": pdf_path.name,
                "Firma": current_org,
                "Lizenzen": licenses,
                "Neue_Lizenzen": new_licenses or 0,
                "Total_Lizenzen": licenses + (new_licenses or 0),
            })

    return rows

def save_as_csv(all_rows: list[dict], date) -> None:
    # In DataFrame überführen
    df = pd.DataFrame(all_rows, columns=["Quelle", "Firma", "Lizenzen", "Neue_Lizenzen", "Total_Lizenzen"])
    if df.empty:
        print("Keine Organisationsdaten gefunden.")
        return
    else:
        df = df.sort_values(by=["Firma"], key=lambda col: col.str.lower())

    # Speichern als CSV
    out_csv = f"{date} ExtractedLicenseOverviewReports.csv"
    df.to_csv(out_csv, index=False, sep=";", encoding="utf-8-sig")
    print(f"\nCSV gespeichert: {out_csv}")

def extract_rows_from_pdfs(pdf_paths):
    all_rows: list[dict] = []
    all_dates: list = []

    for pdf_file in pdf_paths:
        pdf_path = Path(pdf_file)
        try:
            date = get_reporting_interval(pdf_path)
            all_dates.append(date)
            rows = parse_single_pdf(pdf_path)
            all_rows.extend(rows)
            print(f"✓ Verarbeitet: {pdf_path.name}  -> {len(rows)} Einträge")
        except Exception as e:
            print(f"⚠️ Fehler bei {pdf_path.name}: {e}")
            sys.exit(1)

    if not len(set(all_dates)) == 1:
        print(f"⚠️ Fehler, die ausgewählten PDFs sind nicht vom gleichen Zeitraum: {all_dates}")
        sys.exit(1)

    save_as_csv(all_rows, all_dates[0])
    return all_rows, all_dates[0]
