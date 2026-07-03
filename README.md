# Headed Paper PDF Generator

A standalone Windows desktop app that reads a clinic Excel export (`.xls` or
`.xlsx`), overlays each patient's details (name, DOB, MRN, appointment type,
clinic date, and an MRN QR code) onto the `letterhead.pdf` template, and saves
one combined PDF of patient sheets.

## Using the app

1. Run `HeadedPaperPDF.exe` (no Python installation needed — the letterhead
   template is bundled inside the executable).
2. Click **Select Excel File and Generate PDF**.
3. Pick the clinic Excel file, then pick the folder to save the PDF in.
4. The combined PDF is saved as `<clinic date>_OPD_Patient_Sheets.pdf`.

## Getting the executable

Either download the `HeadedPaperPDF-windows` artifact from the
**Build Windows executable** GitHub Actions run, or build it yourself on a
Windows machine:

1. Install Python 3.10+ from python.org (tick "Add to PATH").
2. Double-click `build_windows.bat` (or run `pyinstaller headed_paper.spec`).
3. The app appears at `dist\HeadedPaperPDF.exe`.

## Running from source (any OS)

```
pip install -r requirements.txt
python headed_paper_pdf.py
```

When running from source, `letterhead.pdf` is loaded from the same folder as
the script.
