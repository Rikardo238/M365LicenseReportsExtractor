import sys
import tkinter as tk
from tkinter import filedialog, messagebox

def get_pdf_files():
    """Explorer öffnet sich, wo man die verschiedenen PDFs auswählen kann."""
    root = tk.Tk()
    root.withdraw()

    pdf_files = filedialog.askopenfilenames(
        title="License Overview Reports auswählen (PDF)",
        filetypes=[("PDF-Dateien", "*.pdf")]
    )
    if not pdf_files:
        messagebox.showinfo("Abgebrochen", "Keine PDF-Dateien ausgewählt.")
        sys.exit(1)

    root.destroy()
    return pdf_files

def get_excel_file():
    """Explorer öffnet sich, wo man die verschiedenen PDFs auswählen kann."""
    root = tk.Tk()
    root.withdraw()

    excel_file = filedialog.askopenfilename(
        title="Verrechnung Excel auswählen",
        filetypes=[("Excel-Datei", "*.xlsx")]
    )
    if not excel_file:
        messagebox.showinfo("Abgebrochen", "Keine Excel-Datei ausgewählt.")
        sys.exit(1)

    root.destroy()
    return excel_file
