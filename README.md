Below is an updated `README.md` that **focuses on the GTK for Windows Runtime approach** (rather than MSYS2) and mentions the **`tmp_attachments`** directory for processing attachments.

```markdown
# EML to PDF Converter

This tool converts `.eml` (email) files to PDF on **Windows**. It:

- Renders the email body (HTML preferred, or plain text).
- Embeds inline images (`cid:` references).
- Appends PDF attachments (each prefixed by a small title page).
- Stores attachments in a local folder named **`tmp_attachments`** during processing.

---

## 1. Installation (Windows + GTK Runtime)

### 1.1 Prerequisites

1. **Python 3.9+**  
   - Install from [python.org](https://www.python.org/downloads/) or the Microsoft Store.  
   - Verify with:
     ```powershell
     python --version
     ```

2. **GTK for Windows Runtime**  
   - Download from the [GTK for Windows Runtime Installer Releases](https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases).  
   - Run the `.exe` installer and **check** “Set up PATH environment variable.”  
   - Close and reopen your terminal (so the new PATH takes effect).

### 1.2 Create a Python Virtual Environment

In **PowerShell** (example):

```powershell
# Navigate to your project folder
cd "C:\path\to\the\project"

# Create a new virtual environment
python -m venv venv

# Activate the venv
.\venv\Scripts\activate
```

(On Linux/macOS, you’d do `python3 -m venv venv && source venv/bin/activate`, but this doc focuses on Windows.)

### 1.3 Install Python Dependencies

Inside your **activated** virtual environment

```powershell
pip install --upgrade pip
pip install weasyprint PyPDF2
```

Verify WeasyPrint:

```powershell
python -c "import weasyprint; print('WeasyPrint OK')"
```

---

## 2. Usage

With your **venv** activated, run the script:

```powershell
python eml2pdf.py "C:\path\to\eml_files" "C:\path\to\output_pdfs"
```

- **`"C:\path\to\eml_files"`**: Directory containing `.eml` files to convert.
- **`"C:\path\to\output_pdfs"`**: Destination for the resulting `.pdf` files.

### Temporary Attachments Folder

The script processes attachments by **extracting them** into a directory called **`tmp_attachments`**. It then merges these attachments (if PDFs) into the final PDF, and cleans up temporary files afterward.

Feel free to customize or change the `tmp_attachments` path in the script as needed.

---

## 3. Handling Corrupted PDFs

If an **attached PDF** is **corrupted** or in an unexpected format, PyPDF2 may raise an error like:

```
PyPDF2.errors.PdfReadError: Invalid Elementary Object ...
```

By default, this halts processing. If you prefer:

- **Skip** unreadable PDF attachments, or
- **Insert** a “Broken PDF” notice page,

you can modify the script’s `try/except` logic around the PyPDF2 `append()` call.

---

## 4. Project Structure

A typical layout:

```
eml2pdf.py
venv/
    (your virtual environment)
tmp_attachments/
    (created automatically for attachments)
requirements.txt
README.md
```

**`eml2pdf.py`** is your main script for:
1. Parsing `.eml` files
2. Embedding inline images
3. Appending PDF attachments (stored temporarily in `tmp_attachments`)
4. Generating final PDFs

---

## 5. License

(C) 2025 **Rafael Borja**  
Licensed under [Apache 2.0](LICENSE).

You’re free to copy, modify, and distribute under these terms.

---

## 6. Troubleshooting

- **Missing DLL error (`gobject-2.0-0`, etc.)**:
    - Confirm GTK for Windows is installed and you checked the box to update `PATH`.
- **WeasyPrint Warnings** (`GLib-GIO-WARNING: Unexpectedly, UWP app…`):
    - Often harmless. You can ignore them, or try setting environment variables (e.g. `G_MESSAGES_DEBUG=none`).
- **PermissionError** when deleting temp files on Windows:
    - Ensure you only remove them after PyPDF2 finishes merging (i.e., after `merger.close()`).
