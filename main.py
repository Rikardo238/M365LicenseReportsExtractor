# Glaronia Informatik AG
# Rikardo Stoilov
# 22.08.2024

# Python 3.13
# Benötigt: pip install pdfplumber pandas pyinstaller

# Mit diesem Befehl im Terminal, kann das main.py in einem .exe umgewandelt werden.
# pyinstaller --onefile --icon=logo.ico --name License_Reports_Extractor --clean main.py

import os
import selector
import m365_extractor
import cloud_connect_extractor
import excel_creator

def main():
    try:
        pdf_files = selector.get_pdf_files()

        # Unterscheidung
        if len(pdf_files) > 1:
            # Prüfen, ob ALLE Dateien die Wörter license + overview + report enthalten
            if all(all(word in os.path.basename(f).lower() for word in ["license", "overview", "report"]) for f in pdf_files):
                print("M365 License Usage")
                rows, date = m365_extractor.extract_rows_from_pdfs(pdf_files)
                excel_creator.save_as_excel(rows, date, "M365 License Reports Extracted")
        elif len(pdf_files) == 1:
            # Prüfen, ob die Wörter license + usage + report enthalten sind
            if all(word in os.path.basename(pdf_files[0]).lower() for word in ["license", "usage", "report"]):
                print("Cloud Connect License Usage")
                rows, date = cloud_connect_extractor.extract_rows_from_pdf(pdf_files[0])
                excel_creator.save_as_excel(rows, date, "Cloud Connect License Reports Extracted")
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")
    finally:
        input("Drücke Enter um die Konsole zu schliessen...")

if __name__ == "__main__":
    main()
