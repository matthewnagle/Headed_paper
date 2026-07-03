"""Headed Paper PDF Generator.

Tkinter desktop app that reads a clinic Excel export, overlays patient
details onto the letterhead template, and saves one combined PDF.

Packaged as a standalone Windows executable with PyInstaller (see
build_windows.bat / headed_paper.spec). The letterhead.pdf template is
bundled inside the executable.
"""

import io
import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

import fitz  # PyMuPDF
import pandas as pd
import qrcode
import xlrd

APP_TITLE = "PDF Form Generator"
TEMPLATE_FILENAME = "letterhead.pdf"


def resource_path(filename):
    """Resolve a bundled resource both in dev and inside a PyInstaller exe."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, filename)


def find_template():
    """Locate letterhead.pdf: bundled copy first, then next to the exe/script."""
    candidates = [
        resource_path(TEMPLATE_FILENAME),
        os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), TEMPLATE_FILENAME),
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    # Fall back to asking the user where the template lives.
    return filedialog.askopenfilename(
        title="Locate letterhead.pdf template",
        filetypes=(("PDF files", "*.pdf"), ("All files", "*.*")),
    )


def process_excel_file(file_path):
    if file_path.lower().endswith(".xls"):
        book = xlrd.open_workbook_xls(file_path, ignore_workbook_corruption=True)
        df = pd.read_excel(book, skiprows=4, usecols=[4, 6, 8, 10])
    else:
        df = pd.read_excel(file_path, skiprows=4, usecols=[4, 6, 8, 10])
    return df.fillna("")


def extract_clinic_date(file_path):
    if file_path.lower().endswith(".xls"):
        workbook = xlrd.open_workbook_xls(file_path, ignore_workbook_corruption=True)
        value = workbook.sheet_by_index(0).cell_value(2, 1)
    else:
        header = pd.read_excel(file_path, header=None, nrows=3)
        value = header.iloc[2, 1]
    return str(value)[-10:]


def generate_pdf(file_path, pdf_template_path, output_directory):
    df = process_excel_file(file_path)

    clinic_date = extract_clinic_date(file_path)
    # Dates like 03/07/2026 contain characters that are invalid in file names.
    safe_date = re.sub(r'[\\/:*?"<>|]', "-", clinic_date)
    combined_pdf_path = os.path.join(
        output_directory, f"{safe_date}_OPD_Patient_Sheets.pdf"
    )

    # Coordinates for fields
    name_coords = (70, 250)
    DOB_coords = (70, 273)
    MRN_coords = (70, 296)
    date_coords = (450, 200)
    app_type_coords = (450, 215)
    medicolegal_coords = {
        "OCCUPATION": (70, 312),
        "DATE OF INJURY": (70, 324),
        "DATE OF EXAMINATION": (70, 336),
        "HISTORY/INJURIES": (70, 358),
        "TREATMENT": (70, 480),
        "PREVIOUS HISTORY": (70, 550),
        "PRESENT COMPLAINTS": (70, 590),
        "WORK": (70, 660),
        "EXAMINATION": (70, 690),
        "OPINION": (70, 750),
    }
    qr_code_coords = fitz.Rect(65, 190, 115, 240)

    combined_pdf = fitz.open()

    for index, row in df.iterrows():
        pdf_document = fitz.open(pdf_template_path)

        name = str(row.get("Patient Name", ""))
        dob = str(row.get("DOB", ""))
        mrn = str(row.get("MRN1", ""))
        app_type = str(row.get("Description", ""))

        page = pdf_document[0]
        page.insert_text(name_coords, name, fontsize=10, color=(0, 0, 0))
        page.insert_text(DOB_coords, dob, fontsize=10, color=(0, 0, 0))
        page.insert_text(MRN_coords, mrn, fontsize=10, color=(0, 0, 0))
        page.insert_text(date_coords, clinic_date, fontsize=10, color=(0, 0, 0))
        page.insert_text(app_type_coords, app_type, fontsize=10, color=(0, 0, 0))

        if app_type == "Medicolegal":
            for key, coords in medicolegal_coords.items():
                page.insert_text(coords, key + ":", fontsize=10, color=(0, 0, 0))

        if mrn:
            qr_image = qrcode.make(mrn).get_image().resize((100, 100))
            buffer = io.BytesIO()
            qr_image.save(buffer, format="PNG")
            page.insert_image(qr_code_coords, stream=buffer.getvalue())

        combined_pdf.insert_pdf(pdf_document)
        pdf_document.close()

    combined_pdf.save(combined_pdf_path)
    combined_pdf.close()
    return combined_pdf_path


def select_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xls *.xlsx"), ("All files", "*.*")),
    )
    if not file_path:
        return

    output_directory = filedialog.askdirectory(title="Select Folder to Save PDF")
    if not output_directory:
        return

    pdf_template_path = find_template()
    if not pdf_template_path:
        messagebox.showerror(
            "Error", "letterhead.pdf template not found. Cannot generate PDF."
        )
        return

    try:
        combined_pdf_path = generate_pdf(file_path, pdf_template_path, output_directory)
        messagebox.showinfo(
            "Success",
            f"Combined PDF file created successfully! Saved as {combined_pdf_path}",
        )
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def main():
    app = tk.Tk()
    app.title(APP_TITLE)
    app.geometry("360x120")
    app.resizable(False, False)

    select_file_button = tk.Button(
        app,
        text="Select Excel File and Generate PDF",
        command=select_file,
        padx=12,
        pady=8,
    )
    select_file_button.pack(expand=True, pady=20)

    app.mainloop()


if __name__ == "__main__":
    main()
