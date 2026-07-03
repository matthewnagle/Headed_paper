#code to for TKinter app to gerneate pdf 

#code where user selects output location
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import xlrd
import fitz  # PyMuPDF
import qrcode
import tempfile
from datetime import datetime
import os

# Function to process the Excel file and create a DataFrame
def process_excel_file(file_path):
    df = pd.read_excel(
        xlrd.open_workbook_xls(file_path, ignore_workbook_corruption=True),
        skiprows=4,
        usecols=[4, 6, 8, 10]
    )
    df = df.fillna('')  # Replace NaN with empty strings
    return df

# Function to extract the clinic date
def extract_clinic_date(file_path):
    workbook = xlrd.open_workbook(file_path, ignore_workbook_corruption=True)
    sheet = workbook.sheet_by_index(0)
    clinic_date = str(sheet.cell_value(2, 1))[-10:]
    return clinic_date

# Function to generate the combined PDF
def generate_pdf(file_path, pdf_template_path, output_directory):
    # Process the Excel file
    df = process_excel_file(file_path)

    # Extract the clinic date for naming the output file
    clinic_date = extract_clinic_date(file_path)
    combined_pdf_path = os.path.join(output_directory, f"{clinic_date}_OPD_Patient_Sheets.pdf")

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

    # Create a new empty PDF to hold all combined pages
    combined_pdf = fitz.open()

    for index, row in df.iterrows():
        # Open the PDF template for each row
        pdf_document = fitz.open(pdf_template_path)
        
        name = str(row.get('Patient Name', ''))
        dob = str(row.get('DOB', ''))
        mrn = str(row.get('MRN1', ''))
        app_type = str(row.get('Description', ''))
        
        # Access the first page of the template
        page = pdf_document[0]
        page.insert_text(name_coords, name, fontsize=10, color=(0, 0, 0))
        page.insert_text(DOB_coords, dob, fontsize=10, color=(0, 0, 0))
        page.insert_text(MRN_coords, mrn, fontsize=10, color=(0, 0, 0))
        page.insert_text(date_coords, clinic_date, fontsize=10, color=(0, 0, 0))
        page.insert_text(app_type_coords, app_type, fontsize=10, color=(0, 0, 0))
        
        # Add additional fields for Medicolegal cases
        if app_type == 'Medicolegal':
            for key, coords in medicolegal_coords.items():
                page.insert_text(coords, key + ':', fontsize=10, color=(0, 0, 0))

        # Generate and place QR code for MRN if it has a value
        if mrn:
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as temp_file:
                qr = qrcode.make(mrn)
                qr_image = qr.resize((100, 100))
                qr_image.save(temp_file, format="PNG")
                temp_file_path = temp_file.name
            
            page.insert_image(qr_code_coords, filename=temp_file_path)
            os.remove(temp_file_path)

        # Add the filled page to the combined PDF
        combined_pdf.insert_pdf(pdf_document)
        pdf_document.close()

    # Save the combined PDF
    combined_pdf.save(combined_pdf_path)
    combined_pdf.close()
    messagebox.showinfo("Success", f"Combined PDF file created successfully! Saved as {combined_pdf_path}")

# Function to handle file selection and PDF generation
def select_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=(("Excel files", "*.xls *.xlsx"), ("All files", "*.*"))
    )
    if not file_path:
        return

    output_directory = filedialog.askdirectory(
        title="Select Folder to Save PDF"
    )
    if not output_directory:
        return

    try:
        pdf_template_path = '/Users/matthewnagle/Dropbox/Work/Rooms/headed_paper_app/letterhead.pdf'  # Replace with actual PDF template path
        generate_pdf(file_path, pdf_template_path, output_directory)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Set up the Tkinter app
app = tk.Tk()
app.title("PDF Form Generator")

# Button to select file and generate PDF
select_file_button = tk.Button(app, text="Select Excel File and Generate PDF", command=select_file)
select_file_button.pack(pady=20)

# Run the Tkinter app
app.mainloop()
