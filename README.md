Below is an updated **`eml2pdf.py`** script that addresses a scenario where **some PDF attachments fail** but **don’t raise** a `PyPDF2.errors.PdfReadError`. For instance, PyPDF2 might throw different exceptions or the attachment might be missing critical data but not produce the expected `PdfReadError`. In these cases, you previously saw a warning in logs but **the summary** incorrectly showed zero broken attachments.

### How We Fixed It

1. **Catch All Exceptions** When Merging an Attachment  
   Previously, only `PdfReadError` exceptions were caught. Now we **catch any exception** from PyPDF2 and mark the attachment as broken, ensuring it appears in the final report.

2. **Log** Exactly Which Exception Occurred  
   If it’s **not** a `PdfReadError`, we log a generic `[WARNING]` message indicating an “unknown error.”

3. **Ensure** `broken_pdfs` is appended in both cases, guaranteeing that the final summary includes these attachments.

With these changes, your **report** should correctly list all problematic attachments, whether they fail with `PdfReadError` or some other exception.

---

## Updated `eml2pdf.py`

```python
#!/usr/bin/env python3
"""
EML to PDF Converter
(C) 2025 Rafael Borja

Features:
 - Convert .eml files to PDF (inline images + PDF attachments).
 - Email headers (Date, From, To, CC, BCC, Subject) shown at the top of the PDF.
 - Stores attachments and temporary PDFs in ./tmp_attachments.
 - Corrupted or unreadable PDF attachments get a "Broken PDF" notice instead of crashing.
 - Logs progress at each step.
 - Can optionally merge all resulting PDFs into one file (via --merge).
 - Can optionally create a text report of the summary in output_dir (via --report).

Usage:
  python eml2pdf.py <input_eml_directory> <output_pdf_directory> [--merge] [--report]
"""

import os
import base64
import email
import re
import uuid
from email.policy import default
from email.header import decode_header
from pathlib import Path
import argparse

# Attempt to hide GLib-GIO "Unexpectedly, UWP app..." warnings on Windows
os.environ['G_MESSAGES_DEBUG'] = 'none'

import weasyprint
import PyPDF2


def log(msg: str):
    """Simple logger function."""
    print(f"[INFO] {msg}")


def decode_str(s):
    """
    Decode MIME-encoded strings, e.g. '=?utf-8?...?='.
    Returns a decoded Unicode string.
    """
    if not s:
        return ""
    parts = decode_header(s)
    decoded = ""
    for text, enc in parts:
        if enc:
            decoded += text.decode(enc, errors='ignore')
        elif isinstance(text, bytes):
            decoded += text.decode('utf-8', errors='ignore')
        else:
            decoded += text
    return decoded


def make_attachment_title_pdf(attachment_name: str) -> bytes:
    """
    Create a single-page PDF in memory with a header "Attachment: <filename>".
    Returns the PDF bytes.
    """
    html = f"""
    <html>
      <body>
        <h1 style="margin-top: 100px; text-align:center;">
          Attachment: {attachment_name}
        </h1>
      </body>
    </html>
    """
    return weasyprint.HTML(string=html).write_pdf()


def make_broken_pdf_notice(filename: str) -> bytes:
    """
    Create a single-page PDF in memory indicating the attachment was unreadable.
    Returns the PDF bytes.
    """
    html = f"""
    <html>
      <body>
        <h2 style="color:red; text-align:center; margin-top:100px;">
          Failed to Merge Attachment
        </h2>
        <p style="text-align:center;">{filename}</p>
        <p style="text-align:center;">
          This PDF may be corrupted or use unsupported features.
        </p>
      </body>
    </html>
    """
    return weasyprint.HTML(string=html).write_pdf()


def extract_email_headers(msg):
    """
    Extract key headers (Date, From, To, CC, BCC, Subject) from the email.Message object.
    Return them as a dict with decoded strings.
    Note: BCC is often missing in raw EML unless specifically saved that way.
    """
    date_raw = msg["Date"] or ""
    from_raw = msg["From"] or ""
    to_raw = msg["To"] or ""
    cc_raw = msg["Cc"] or ""
    bcc_raw = msg["Bcc"] or ""
    subj_raw = msg["Subject"] or ""

    return {
        "date": decode_str(date_raw),
        "from": decode_str(from_raw),
        "to": decode_str(to_raw),
        "cc": decode_str(cc_raw),
        "bcc": decode_str(bcc_raw),
        "subject": decode_str(subj_raw),
    }


def make_header_html(headers: dict) -> str:
    """
    Given a dict with 'date', 'from', 'to', 'cc', 'bcc', 'subject',
    build a small HTML snippet for display at the top.
    """
    return f"""
<div style="font-family: sans-serif; border:1px solid #ccc; padding:10px; margin-bottom:10px;">
  <p><strong>Date:</strong> {headers["date"]}</p>
  <p><strong>From:</strong> {headers["from"]}</p>
  <p><strong>To:</strong> {headers["to"]}</p>
  <p><strong>CC:</strong> {headers["cc"]}</p>
  <p><strong>BCC:</strong> {headers["bcc"]}</p>
  <p><strong>Subject:</strong> {headers["subject"]}</p>
</div>
"""


def extract_html_and_inline_images(msg):
    """
    Extract the best available HTML (or text) from an email and find attachments.

    Returns:
      - html_content (string): The HTML body
      - attachments (list): A list of email parts considered attachments
    """
    html_content = None
    text_content = None
    attachments = []

    if msg.is_multipart():
        # Traverse all parts
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get("Content-Disposition") or "").lower()
            content_id = part.get("Content-ID")

            if content_type == "text/html" and "attachment" not in content_disposition:
                # Prefer HTML over plain text
                html_payload = part.get_payload(decode=True)
                html_charset = part.get_content_charset() or "utf-8"
                html_content = html_payload.decode(html_charset, errors="ignore")

            elif content_type == "text/plain" and "attachment" not in content_disposition:
                # Fallback to text
                text_payload = part.get_payload(decode=True)
                text_charset = part.get_content_charset() or "utf-8"
                text_content = text_payload.decode(text_charset, errors="ignore")

            else:
                # Possibly an attachment
                attachments.append(part)
    else:
        # Single-part message
        content_type = msg.get_content_type()
        if content_type == "text/html":
            html_payload = msg.get_payload(decode=True)
            html_charset = msg.get_content_charset() or "utf-8"
            html_content = html_payload.decode(html_charset, errors="ignore")
        elif content_type == "text/plain":
            text_payload = msg.get_payload(decode=True)
            text_charset = msg.get_content_charset() or "utf-8"
            text_content = text_payload.decode(text_charset, errors="ignore")

    # If we have no HTML, wrap plain text in minimal HTML
    if not html_content:
        if text_content:
            html_content = f"<html><body><pre>{text_content}</pre></body></html>"
        else:
            html_content = "<html><body>(No content)</body></html>"

    # Embed inline images (cid:...)
    for part in attachments:
        content_id = part.get("Content-ID")
        if content_id:
            cid_clean = content_id.strip("<>")
            if cid_clean in html_content:
                payload = part.get_payload(decode=True)
                maintype = part.get_content_maintype()
                subtype = part.get_content_subtype()
                base64_data = base64.b64encode(payload).decode("utf-8")
                data_uri = f"data:{maintype}/{subtype};base64,{base64_data}"

                pattern = re.compile(r'src=["\']cid:' + re.escape(cid_clean) + r'["\']')
                replacement = f'src="{data_uri}"'
                html_content = pattern.sub(replacement, html_content)

    return html_content, attachments


def create_attachment_list_html(attachments):
    """
    Create an HTML snippet listing non-inline attachments at the end of the email.
    """
    attach_names = []
    for part in attachments:
        filename = decode_str(part.get_filename() or "")
        if not filename:
            continue
        content_disposition = str(part.get("Content-Disposition") or "").lower()
        content_id = part.get("Content-ID")
        # We'll say it's a user-facing attachment if it has a filename and is not purely inline
        if "attachment" in content_disposition or not content_id:
            attach_names.append(filename)

    if not attach_names:
        return ""

    items = "".join(f"<li>{fn}</li>" for fn in attach_names)
    return f"""
    <hr>
    <h3>Attachments in this email:</h3>
    <ul>
      {items}
    </ul>
    """


def convert_html_to_pdf(html_content, output_pdf_path):
    """
    Convert the given HTML string to PDF using WeasyPrint, writing the result to output_pdf_path.
    """
    pdf_bytes = weasyprint.HTML(string=html_content).write_pdf()
    with open(output_pdf_path, 'wb') as f:
        f.write(pdf_bytes)


def append_pdf_attachments_to_pdf(attachments, base_pdf_path):
    """
    Append PDF attachments to the base PDF (email body).
     - For each PDF, insert a one-page "Attachment: <filename>" PDF.
     - If merging fails, insert a "Broken PDF" notice instead.
     - Return a tuple: (# of PDF attachments processed, list of broken PDF filenames).
    """
    merger = PyPDF2.PdfMerger()
    merger.append(base_pdf_path)

    # Ensure tmp_attachments folder
    tmp_dir = Path("tmp_attachments")
    tmp_dir.mkdir(exist_ok=True)

    temp_files = []
    pdf_count = 0
    broken_pdfs = []

    for part in attachments:
        filename = decode_str(part.get_filename() or "")
        if filename.lower().endswith(".pdf"):
            pdf_count += 1
            log(f"  -> Processing PDF attachment: {filename}")

            # 1) Title page
            attachment_title_bytes = make_attachment_title_pdf(filename)
            title_pdf_path = tmp_dir / f"title_{uuid.uuid4()}.pdf"
            with open(title_pdf_path, 'wb') as tf:
                tf.write(attachment_title_bytes)
            temp_files.append(title_pdf_path)

            merger.append(str(title_pdf_path))

            # 2) Write the actual attachment to disk
            payload = part.get_payload(decode=True)
            pdf_temp_path = tmp_dir / f"attach_{uuid.uuid4()}.pdf"
            with open(pdf_temp_path, 'wb') as f:
                f.write(payload)
            temp_files.append(pdf_temp_path)

            # Attempt to merge the PDF
            try:
                merger.append(str(pdf_temp_path))
            except PyPDF2.errors.PdfReadError as e:
                # Mark this attachment as broken (PyPDF2 recognized the data as invalid PDF)
                log(f"[WARNING] Could not merge PDF '{filename}' (PdfReadError): {e}")
                broken_pdfs.append(filename)

                broken_notice_bytes = make_broken_pdf_notice(filename)
                broken_notice_path = tmp_dir / f"broken_{uuid.uuid4()}.pdf"
                with open(broken_notice_path, 'wb') as bf:
                    bf.write(broken_notice_bytes)
                temp_files.append(broken_notice_path)

                # Merge the broken notice
                merger.append(str(broken_notice_path))

            except Exception as ex:
                # Catch any other exception (e.g., unsupported PDF features, I/O issues)
                log(f"[WARNING] Could not merge PDF '{filename}' (unknown error): {ex}")
                broken_pdfs.append(filename)

                broken_notice_bytes = make_broken_pdf_notice(filename)
                broken_notice_path = tmp_dir / f"broken_{uuid.uuid4()}.pdf"
                with open(broken_notice_path, 'wb') as bf:
                    bf.write(broken_notice_bytes)
                temp_files.append(broken_notice_path)

                # Merge the broken notice
                merger.append(str(broken_notice_path))

    merger.write(base_pdf_path)
    merger.close()

    # Cleanup temp files
    for tmpf in temp_files:
        try:
            os.remove(tmpf)
        except OSError as e:
            log(f"[WARNING] Could not remove temp file {tmpf}: {e}")

    return pdf_count, broken_pdfs


def eml_to_pdf(eml_path, output_pdf_path):
    """
    Main routine that:
     1) Parses .eml for headers + HTML/text & attachments
     2) Adds the header block, then an attachment list at the end
     3) Converts email body to PDF
     4) Merges PDF attachments
     5) Returns (# of PDF attachments processed, list_of_broken_pdfs).
    """
    log(f"Reading EML: {eml_path}")
    with open(eml_path, 'rb') as f:
        raw_data = f.read()

    msg = email.message_from_bytes(raw_data, policy=default)

    # Extract headers
    hdrs = extract_email_headers(msg)

    # Extract body & attachments
    html_content, attachments = extract_html_and_inline_images(msg)

    # Insert the header at the top
    header_html = make_header_html(hdrs)
    if "<body>" in html_content:
        html_content = html_content.replace("<body>", f"<body>{header_html}", 1)
    else:
        # If no <body> tag found, wrap everything
        html_content = f"<html><body>{header_html}{html_content}</body></html>"

    # Insert an attachment list at the bottom
    html_content += create_attachment_list_html(attachments)

    # Convert body to PDF
    log("  -> Converting HTML body to PDF")
    convert_html_to_pdf(html_content, output_pdf_path)

    # Merge PDF attachments
    log("  -> Checking for PDF attachments to merge")
    pdf_count, broken_list = append_pdf_attachments_to_pdf(attachments, output_pdf_path)
    return pdf_count, broken_list


def convert_eml_files_in_directory(eml_dir, output_dir, merge_all=False, report=False):
    """
    Processes all .eml files in `eml_dir` and writes .pdf files to `output_dir`.
    If merge_all=True, merges all resulting PDFs into a single "merged_all.pdf".
    If report=True, writes a 'eml2pdf.txt' file in `output_dir` containing the summary.
    """
    entries = [e for e in os.scandir(eml_dir) if e.is_file() and e.name.lower().endswith(".eml")]
    total_files = len(entries)
    if total_files == 0:
        log("No EML files found.")
        return

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    # Tracking summary info
    total_pdf_attachments = 0
    broken_pdfs = []  # list of (filename) that failed
    generated_pdfs = []  # list of all final PDFs for potential merging

    for i, entry in enumerate(entries, start=1):
        log(f"\n=== Processing file {i} of {total_files}: {entry.name} ===")
        output_pdf_path = os.path.join(output_dir, f"{os.path.splitext(entry.name)[0]}.pdf")
        log(f"Output PDF: {output_pdf_path}")

        pdf_count, broken_list = eml_to_pdf(entry.path, output_pdf_path)
        total_pdf_attachments += pdf_count
        broken_pdfs.extend(broken_list)

        # Keep track of newly generated PDF for optional merging
        generated_pdfs.append(output_pdf_path)

    # Merge all PDFs if requested
    merged_filename = None
    if merge_all and generated_pdfs:
        merged_filename = os.path.join(output_dir, "merged_all.pdf")
        log(f"\n[MERGE] Merging all {len(generated_pdfs)} PDFs into {merged_filename}")
        merger = PyPDF2.PdfMerger()
        for pdf_file in generated_pdfs:
            merger.append(pdf_file)
        merger.write(merged_filename)
        merger.close()

    # Prepare summary
    summary_lines = []
    summary_lines.append("\n=== SUMMARY REPORT ===")
    summary_lines.append(f"Total EML files processed: {total_files}")
    summary_lines.append(f"Total PDF attachments processed: {total_pdf_attachments}")

    if broken_pdfs:
        summary_lines.append("Broken/Unreadable PDF attachments:")
        for fn in broken_pdfs:
            summary_lines.append(f"  - {fn}")
    else:
        summary_lines.append("No broken PDF attachments encountered.")

    if merged_filename:
        summary_lines.append(f"Merged all PDFs into: {merged_filename}")

    # Print the summary
    for line in summary_lines:
        log(line)

    # If --report was used, write eml2pdf.txt
    if report:
        report_path = os.path.join(output_dir, "eml2pdf.txt")
        with open(report_path, "w", encoding="utf-8") as rep:
            for line in summary_lines:
                rep.write(line + "\n")
        log(f"[REPORT] Summary written to: {report_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Convert EML files to PDF (headers, inline images, PDF attachments)."
    )
    parser.add_argument("eml_dir", help="Path to the directory containing .eml files.")
    parser.add_argument("output_dir", help="Directory where .pdf files will be saved.")
    parser.add_argument("--merge", action="store_true",
                        help="Merge all generated PDFs into one 'merged_all.pdf' at the end.")
    parser.add_argument("--report", action="store_true",
                        help="Generate a text file 'eml2pdf.txt' with the summary report.")

    args = parser.parse_args()

    convert_eml_files_in_directory(
        eml_dir=args.eml_dir,
        output_dir=args.output_dir,
        merge_all=args.merge,
        report=args.report
    )


if __name__ == "__main__":
    main()
```

---

## Updated **README.md**

Below is an updated example **README.md** reflecting the fix for attachments that fail under exceptions other than `PdfReadError`, plus usage instructions:

```markdown
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
```

---

### Final Notes

- **Broken Attachments**: Now any attachment that fails for *any reason* gets flagged in `broken_pdfs`, ensuring the summary report is accurate.  
- **WeasyPrint** on Windows: Make sure you install the required **GTK** or **MSYS2** dependencies so WeasyPrint can find Pango/Cairo/GLib.  
- **MIME-Encoded Headers**: The script decodes these using Python’s `decode_header`, so accented subjects, etc. show up properly.