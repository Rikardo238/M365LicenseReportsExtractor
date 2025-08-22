import sys
import tkinter as tk
from tkinter import filedialog, messagebox

def get_pdf_files():
    """Explorer öffnet sich, wo man die verschiedenen PDFs auswählen kann."""
    root = tk.Tk()
    root.withdraw()

    pdf_files = filedialog.askopenfilenames(
        title="License Reports auswählen (PDF)",
        filetypes=[("PDF-Dateien", "*.pdf")]
    )
    if not pdf_files:
        messagebox.showinfo("Abgebrochen", "Keine PDF-Dateien ausgewählt.")
        sys.exit(1)

    root.destroy()
    return pdf_files

