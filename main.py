# Glaronia Informatik AG
# Rikardo Stoilov
# 20.08.2024

# Python 3.13
# Benötigt: pip install pdfplumber pandas pyinstaller

# Mit diesem Befehl im Terminal, kann das main.py in einem .exe umgewandelt werden mit dem Namen "LizenzenExtractor".
# pyinstaller --onefile --icon=logo.ico --name LicenseReportsExtractor --clean main.py

import selector
import pdf_extractor

def main():
    try:
        pdf_paths = selector.get_pdf_files()
        pdf_extractor.extract_rows_from_pdfs(pdf_paths)
    except Exception as e:
        print(f"Ein Fehler ist aufgetreten: {e}")
    finally:
        input("Drücke Enter um die Konsole zu schliessen...")

if __name__ == "__main__":
    main()
