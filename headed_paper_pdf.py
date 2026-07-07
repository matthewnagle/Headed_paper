"""Headed Paper PDF Generator.

Tkinter desktop app that reads a clinic Excel export, overlays patient
details onto the letterhead template, and prints the patient sheets.

Workflow: select the Excel file -> the PDF is generated into a temporary
folder -> press Print -> the pages are sent to the chosen printer -> the
app deletes the temporary data and closes. Nothing is left on disk.

Packaged as a standalone Windows executable with PyInstaller (see
build_windows.bat / headed_paper.spec). The letterhead.pdf template is
bundled inside the executable. Printing uses the Windows print spooler
directly (pywin32), so no PDF reader needs to be installed.
"""

import io
import os
import re
import shutil
import sys
import tempfile
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz  # PyMuPDF
import pandas as pd
import qrcode
from PIL import Image

import xlrd

APP_TITLE = "Patient Sheet Printer"
TEMPLATE_FILENAME = "letterhead.pdf"
PRINT_DPI = 300


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


def _xls_cell_value(sheet, row, col, datemode):
    cell = sheet.cell(row, col)
    if cell.ctype == xlrd.XL_CELL_DATE:
        return xlrd.xldate_as_datetime(cell.value, datemode).strftime("%d/%m/%Y")
    if cell.ctype == xlrd.XL_CELL_NUMBER and cell.value == int(cell.value):
        return int(cell.value)  # keep MRNs as 12345678, not 12345678.0
    return cell.value


def process_excel_file(file_path):
    columns = [4, 6, 8, 10]
    if file_path.lower().endswith(".xls"):
        # pandas no longer accepts an xlrd Book object, and going through
        # pandas' own xlrd engine would lose ignore_workbook_corruption,
        # which the clinic exports need - so read the sheet directly.
        book = xlrd.open_workbook_xls(file_path, ignore_workbook_corruption=True)
        sheet = book.sheet_by_index(0)
        headers = [str(sheet.cell_value(4, c)).strip() for c in columns]
        rows = []
        for r in range(5, sheet.nrows):
            values = [_xls_cell_value(sheet, r, c, book.datemode) for c in columns]
            if any(str(v).strip() for v in values):
                rows.append(values)
        df = pd.DataFrame(rows, columns=headers)
    else:
        df = pd.read_excel(file_path, skiprows=4, usecols=columns)
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
    page_count = combined_pdf.page_count
    combined_pdf.close()
    return combined_pdf_path, page_count


def pdf_page_images(pdf_path, dpi=PRINT_DPI):
    """Render each PDF page to a PIL image, ready to send to the printer."""
    doc = fitz.open(pdf_path)
    try:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            yield Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
    finally:
        doc.close()


def list_printers():
    import win32print

    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    printers = [p[2] for p in win32print.EnumPrinters(flags)]
    default = win32print.GetDefaultPrinter()
    return printers, default


def print_pdf(pdf_path, printer_name, job_name="OPD Patient Sheets"):
    """Send every page of the PDF to the printer via the Windows spooler."""
    import win32ui
    from PIL import ImageWin

    hdc = win32ui.CreateDC()
    hdc.CreatePrinterDC(printer_name)
    printable_w = hdc.GetDeviceCaps(8)  # HORZRES
    printable_h = hdc.GetDeviceCaps(10)  # VERTRES

    hdc.StartDoc(job_name)
    try:
        for img in pdf_page_images(pdf_path):
            scale = min(printable_w / img.width, printable_h / img.height)
            w, h = int(img.width * scale), int(img.height * scale)
            x = (printable_w - w) // 2
            y = (printable_h - h) // 2
            hdc.StartPage()
            ImageWin.Dib(img).draw(hdc.GetHandleOutput(), (x, y, x + w, y + h))
            hdc.EndPage()
    except Exception:
        hdc.AbortDoc()
        raise
    else:
        hdc.EndDoc()
    finally:
        hdc.DeleteDC()


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("420x220")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.close)

        self.temp_dir = None
        self.pdf_path = None
        self.page_count = 0

        self.status = tk.StringVar(value="Select the clinic Excel file to begin.")
        tk.Label(self.root, textvariable=self.status, wraplength=380).pack(pady=(16, 8))

        self.select_button = tk.Button(
            self.root, text="1. Select Excel File", command=self.select_file, width=32
        )
        self.select_button.pack(pady=4)

        printer_row = tk.Frame(self.root)
        printer_row.pack(pady=4)
        tk.Label(printer_row, text="Printer:").pack(side=tk.LEFT, padx=(0, 6))
        self.printer_box = ttk.Combobox(printer_row, width=34, state="readonly")
        self.printer_box.pack(side=tk.LEFT)
        self.load_printers()

        self.print_button = tk.Button(
            self.root,
            text="2. Print pages",
            command=self.print_pages,
            width=32,
            state=tk.DISABLED,
        )
        self.print_button.pack(pady=4)

    def load_printers(self):
        try:
            printers, default = list_printers()
            self.printer_box["values"] = printers
            if default in printers:
                self.printer_box.set(default)
            elif printers:
                self.printer_box.set(printers[0])
        except Exception:
            self.printer_box["values"] = []

    def select_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=(("Excel files", "*.xls *.xlsx"), ("All files", "*.*")),
        )
        if not file_path:
            return

        pdf_template_path = find_template()
        if not pdf_template_path:
            messagebox.showerror(
                "Error", "letterhead.pdf template not found. Cannot generate pages."
            )
            return

        self.cleanup_temp()
        try:
            self.temp_dir = tempfile.mkdtemp(prefix="headed_paper_")
            self.pdf_path, self.page_count = generate_pdf(
                file_path, pdf_template_path, self.temp_dir
            )
        except Exception as e:
            self.cleanup_temp()
            messagebox.showerror("Error", f"An error occurred: {e}")
            return

        self.status.set(
            f"{self.page_count} patient sheet(s) ready. "
            "Choose a printer and press Print."
        )
        self.print_button.config(
            text=f"2. Print {self.page_count} page(s)", state=tk.NORMAL
        )

    def print_pages(self):
        printer_name = self.printer_box.get()
        if not printer_name:
            messagebox.showerror("Error", "No printer selected.")
            return

        self.status.set(f"Printing to {printer_name}…")
        self.print_button.config(state=tk.DISABLED)
        self.select_button.config(state=tk.DISABLED)
        self.root.update_idletasks()

        try:
            print_pdf(self.pdf_path, printer_name)
        except Exception as e:
            self.print_button.config(state=tk.NORMAL)
            self.select_button.config(state=tk.NORMAL)
            self.status.set("Printing failed. Try again or pick another printer.")
            messagebox.showerror("Error", f"Printing failed: {e}")
            return

        messagebox.showinfo(
            "Done",
            f"{self.page_count} page(s) sent to {printer_name}. "
            "The app will now close and delete the data.",
        )
        self.close()

    def cleanup_temp(self):
        if self.temp_dir:
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.temp_dir = None
        self.pdf_path = None
        self.page_count = 0

    def close(self):
        self.cleanup_temp()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


def main():
    App().run()


if __name__ == "__main__":
    main()
