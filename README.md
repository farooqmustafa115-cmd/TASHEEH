# Military College Exam Generator

Desktop question-bank and exam-generation app built with Flask and pywebview.

## Portable Folder

Keep `app.exe` and `exam_bank.db` together in the same writable folder. Copy the complete folder to move the application and its questions to another computer. The executable includes the Python dependencies.

## Windows Installer

Run `dist/MilitaryCollegeExam-Setup-x64.exe` and follow the Arabic setup wizard. The installer is for 64-bit Windows 10 or later, bundles the application, its current question database, and the offline Microsoft WebView2 Runtime. It installs WebView2 silently only when it is missing. Python and an internet connection are not required on the target computer.

The installer is per-user and does not require administrator access. The installed database and its `backups` folder are preserved when updating or uninstalling the app.

To rebuild the installer, first build `dist/app.exe` with PyInstaller and place Microsoft's x64 Evergreen Standalone WebView2 Runtime installer at `installer/MicrosoftEdgeWebView2RuntimeInstallerX64.exe`. Download it from [Microsoft LinkId 2124701](https://go.microsoft.com/fwlink/?linkid=2124701), then compile `installer/setup.iss` from PowerShell:

```powershell
Push-Location installer
& "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" setup.iss
Pop-Location
```

## Run from Source

```bash
python -m pip install -r requirements.txt
python app.py
```

The SQLite database is stored beside the running application. Source runs use `exam_bank.db` beside `app.py`; installed and portable executable runs use the database in their application folder.

The exam issuance screen accepts one or more roster Excel files (`.xlsx`, `.xlsm`, or `.xls`).
The selected files are validated together and replace the current roster only after all rows
pass validation; source roster files are not hard-coded in the application. The import ignores
statistical and military-number columns. Each student is identified by their positive sequence
number within their company (from a `التسلسل`/`تسلسل الطالب`/`serial` column, or the row order
within that company when that column is absent). Each student receives one permanent HS256 JWT
identity token without an expiration claim. The token is retained in the main database and in
`qr_database/qr_tokens.db` so booklet issuance and scan correction can resolve the same student
consistently.

Run the focused exam workflow checks with:

```bash
python -m unittest -v test_exam_workflow
```

## Notes

- The app uses a local Flask server on `http://127.0.0.1:54321`.
- Word export requires the `python-docx` package; it is included in the executable build.
- Exam booklet issuance and scan correction use `openpyxl`/`xlrd` for Excel, `qrcode` for booklet QR codes, and PyMuPDF/OpenCV for scanned pages. `pyzbar` is an optional second QR reader, and `pytesseract` is an optional OCR fallback that reads the printed military number when a QR code is unreadable (install Tesseract-OCR or place it in a `Tesseract-OCR` folder beside the app). Install the updated `requirements.txt` before running from source or rebuilding the executable.
- Import student rosters in `.xlsx`, `.xlsm`, or legacy `.xls` format. The app searches the first 20 rows of every worksheet for student name, company, and platoon headers; a sequence column is optional. Leading title rows are allowed. If supplied, each sequence must be a positive, unique number within its company; otherwise the import assigns sequences by the Excel row order within each company. The issuance screen previews every generated booklet before printing. The 30 mm booklet QR is repeated with the printed table header on each booklet page so pages remain identifiable during scanning. Scan booklets into one PDF or as image files in any common format (PNG, JPG, JPEG, BMP, TIFF, WEBP, GIF); multiple files can be selected at once, and pages of each booklet must keep their order. Automatic bubble reading reports unclear pages/questions for human review; verify sample scans from the actual scanner and printer before relying on final grades.
- The folder must be writable so the app can save questions and create backups.
