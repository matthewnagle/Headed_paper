# Patient Sheet Printer

A standalone Windows desktop app that reads a clinic Excel export (`.xls` or
`.xlsx`), overlays each patient's details (name, DOB, MRN, appointment type,
clinic date, and an MRN QR code) onto the `letterhead.pdf` template, and
prints the patient sheets. Nothing is saved: the pages are built in a
temporary folder, sent straight to the printer, and deleted when the app
closes.

## Using the app

1. Run `HeadedPaperPDF.exe` (no Python or PDF reader needed — the letterhead
   template is bundled inside the executable, and printing goes directly to
   the Windows print spooler).
2. Click **1. Select Excel File** and pick the clinic Excel file. The app
   builds the patient sheets and shows how many pages are ready.
3. Choose a printer (your default printer is pre-selected).
4. Click **2. Print pages**. Once the pages are sent to the printer, the app
   deletes the temporary data and closes itself.

Closing the window at any point also deletes any generated data.

## Getting the executable

Either download the `HeadedPaperPDF-windows` artifact from the
**Build Windows executable** GitHub Actions run, or build it yourself on a
Windows machine:

1. Install Python 3.10+ from python.org (tick "Add to PATH").
2. Double-click `build_windows.bat` (or run `pyinstaller headed_paper.spec`).
3. The app appears at `dist\HeadedPaperPDF.exe`.

## Running from source

```
pip install -r requirements.txt
python headed_paper_pdf.py
```

When running from source, `letterhead.pdf` is loaded from the same folder as
the script. Printing requires Windows (pywin32); the PDF generation itself
works on any OS.
