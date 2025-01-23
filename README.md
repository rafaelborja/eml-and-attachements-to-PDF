# EML to PDF Converter

A Python command-line tool to convert `.eml` (email) files to PDF on **Windows or other platforms**. This project:

- **Extracts HTML or text** from each `.eml` email.
- **Embeds inline images** (`cid:` references) by base64-encoding them in the PDF.
- **Appends PDF attachments** (with a title page) to each email’s PDF.
- **Displays email headers** (Date, From, To, CC, BCC, Subject) at the top of the PDF.
- **Handles corrupted/unreadable PDFs** by inserting a “Broken PDF” notice page, rather than crashing.
- Optionally **merges** all produced PDFs into a single file called `merged_all.pdf`.
- Optionally **generates a text-based summary report** `eml2pdf.txt`.

## 1. Installation

1. **Install Python 3.9+**. On Windows, also install the dependencies for [WeasyPrint](https://weasyprint.org/) (GTK for Windows Runtime or MSYS2).  
2. Create and activate a **virtual environment** (optional, but recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # Linux/macOS
   # or .\venv\Scripts\activate in PowerShell on Windows
   ```
3. **Install** the required Python libraries:
   ```bash
   pip install weasyprint PyPDF2
   ```
4. **Verify** WeasyPrint can run:
   ```bash
   python -c "import weasyprint; print('WeasyPrint OK!')"
   ```

## 2. Usage

```bash
python eml2pdf.py <EML_DIRECTORY> <OUTPUT_PDF_DIRECTORY> [--merge] [--report]
```

Examples:
```bash
# Basic usage (convert each EML to an individual PDF)
python eml2pdf.py ./eml_files ./pdf_output

# Merge all PDFs into a single file
python eml2pdf.py ./eml_files ./pdf_output --merge

# Generate a text report summarizing all results
python eml2pdf.py ./eml_files ./pdf_output --report

# Do both
python eml2pdf.py ./eml_files ./pdf_output --merge --report
```

### Output

- **One `.pdf`** per `.eml` in the specified `<OUTPUT_PDF_DIRECTORY>`.
- If `--merge` is used, a **`merged_all.pdf`** containing all generated PDFs.
- If `--report` is used, a **`eml2pdf.txt`** text file with a final summary (e.g., total `.eml` files, how many PDF attachments were processed, which attachments were broken).

## 3. Handling Broken Attachments

Some PDF attachments may be corrupted or use features that PyPDF2 can’t parse. This script:
- **Catches all exceptions** when merging each PDF attachment.
- Logs a warning and inserts a “Broken PDF” notice page for that attachment.
- Records the filename in a broken attachments list, visible in the final summary.

## 4. Project Structure

A typical layout:
```
eml2pdf.py
tmp_attachments/
venv/
requirements.txt
README.md
```
Where **`tmp_attachments/`** is automatically created for storing temporary PDFs (title pages, broken notices, etc.) during processing.

## 5. License

(C) 2025 Rafael Borja  
Distributed under the [Apache 2.0 License](LICENSE).

---


### Final Notes

- **Broken Attachments**: Now any attachment that fails for *any reason* gets flagged in `broken_pdfs`, ensuring the summary report is accurate.  
- **WeasyPrint** on Windows: Make sure you install the required **GTK** or **MSYS2** dependencies so WeasyPrint can find Pango/Cairo/GLib.  
- **MIME-Encoded Headers**: The script decodes these using Python’s `decode_header`, so accented subjects, etc. show up properly.