# Glaronia Informatik AG
# Rikardo Stoilov
# 15.08.2024

# Python 3.13
# Benötigt: pip install pdfplumber pandas pyinstaller

# Mit diesem Befehl im Terminal, kann das main.py in einem .exe umgewandelt werden mit dem Namen "LizenzenExtractor".
# pyinstaller --onefile --icon=logo.ico --name LicenseReportsExtractor --clean main.py

import re
import pdfplumber
import pandas as pd
from pathlib import Path

# --- Muster (ignoriert Gross-/Kleinschreibung) ---
ORG_PAT = re.compile(r"Organization:\s*(.+)", re.IGNORECASE)
TOTAL_PAT = re.compile(r"Total number of licenses:\s*(\d+)", re.IGNORECASE)
NEW_PAT = re.compile(r"Total number of new licenses:\s*(\d+)", re.IGNORECASE)

def parse_single_pdf(pdf_path: Path) -> list[dict]:
    """Liest EIN PDF, überspringt die erste Seite und gibt eine Liste mit dicts je Organisation zurück."""
    rows = []
    current_org = None
    total_licenses = None
    new_licenses = None

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages[1:]:  # Seite 0 überspringen
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
                            "Lizenzen_gesamt": total_licenses,
                            "Neue_Lizenzen": new_licenses or 0
                        })
                    current_org = org_match.group(1).strip()
                    total_licenses = None
                    new_licenses = None
                    continue

                # Gesamt-Lizenzen
                m_total = TOTAL_PAT.search(line)
                if m_total:
                    total_licenses = int(m_total.group(1))
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
                "Lizenzen_gesamt": total_licenses,
                "Neue_Lizenzen": new_licenses or 0
            })

    return rows

def choose_pdfs():
    """Explorer öffnet sich, wo man die verschiedenen PDFs auswählen kann."""
    try:
        from tkinter import Tk, filedialog
        root = Tk()
        root.withdraw()
        file_paths = filedialog.askopenfilenames(
            title="License Overview Reports auswählen (PDF)",
            filetypes=[("PDF-Dateien", "*.pdf")]
        )
        root.update()
        root.destroy()
    except Exception as e:
        print("Konnte den Dateidialog nicht öffnen:", e)
        return e

    if not file_paths:
        print("Keine Dateien ausgewählt.")
        raise Exception

    return file_paths

def main():
    try:
        all_rows: list[dict] = []

        file_paths = choose_pdfs()
        for fp in file_paths:
            pdf_path = Path(fp)
            try:
                rows = parse_single_pdf(pdf_path)
                all_rows.extend(rows)
                print(f"✓ Verarbeitet: {pdf_path.name}  -> {len(rows)} Einträge")
            except Exception as e:
                print(f"⚠️ Fehler bei {pdf_path.name}: {e}")

        # In DataFrame überführen
        df = pd.DataFrame(all_rows, columns=["Quelle", "Firma", "Lizenzen_gesamt", "Neue_Lizenzen"])
        if df.empty:
            print("\nKeine Organisationsdaten gefunden.")
            return
        else:
            df = df.sort_values(by=["Firma"], key=lambda col: col.str.lower())

        # Ausgabe in Konsole
        print("\nErgebnis (alle PDFs zusammen):")
        print(df.to_string(index=False))

        # Speichern als CSV
        out_csv = "Übersicht_Alle_Lizenzen.csv"
        df.to_csv(out_csv, index=False, sep=";", encoding="utf-8-sig")
        print(f"\nCSV gespeichert: {out_csv}")
    except Exception as e:
        return e

if __name__ == "__main__":
    main()
