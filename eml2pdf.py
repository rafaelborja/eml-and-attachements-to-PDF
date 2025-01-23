#!/usr/bin/env python3
"""
EML to PDF Converter
(C) 2025 Rafael Borja

Features:
 - Convert .eml files to PDF (inline images + PDF attachments).
 - Email headers (Date, From, To, CC, BCC, Subject) at the top of the PDF,
   showing both names and addresses.
 - Stores attachments and temporary PDFs in ./tmp_attachments.
 - Handles corrupted/unreadable PDFs by inserting a "Broken PDF" notice.
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
from email.utils import getaddresses
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
    Decode MIME-encoded strings (e.g. '=?utf-8?...?=').
    Returns a Unicode string.
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


def parse_address_header(value: str) -> str:
    """
    Given a header value like 'John Doe <john@example.com>, "Jane" <jane@example.org>',
    return a string with name+email, e.g.:
      'John Doe <john@example.com>, Jane <jane@example.org>'
    """
    if not value:
        return ""
    # getaddresses returns [(name, email), (name, email), ...]
    pairs = getaddresses([value])
    result_list = []
    for (name_raw, email_addr) in pairs:
        # decode the name if it's MIME-encoded
        name_decoded = decode_str(name_raw)
        if name_decoded and email_addr:
            result_list.append(f"{name_decoded} <{email_addr}>")
        elif email_addr:
            # no display name, just email
            result_list.append(email_addr)
        else:
            # fallback, no email
            result_list.append(name_decoded)
    return ", ".join(result_list)


def make_attachment_title_pdf(attachment_name: str) -> bytes:
    """Single-page PDF titled with the attachment's filename."""
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
    """Single-page PDF noting a broken/unreadable attachment."""
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


def extract_email_headers(msg) -> dict:
    """
    Extract key headers (Date, From, To, CC, BCC, Subject),
    parse them into a more readable string with names + addresses.
    """
    def val(header_name):
        return decode_str(msg[header_name] or "")

    from_raw = parse_address_header(val("From"))
    to_raw   = parse_address_header(val("To"))
    cc_raw   = parse_address_header(val("Cc"))
    bcc_raw  = parse_address_header(val("Bcc"))  # Rarely present in raw EML
    subj_raw = val("Subject")
    date_raw = val("Date")

    return {
        "date": date_raw,
        "from": from_raw,
        "to": to_raw,
        "cc": cc_raw,
        "bcc": bcc_raw,
        "subject": subj_raw,
    }


def make_header_html(headers: dict) -> str:
    """
    Given 'date','from','to','cc','bcc','subject', build an HTML snippet.
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
    Extract HTML or fallback to text. Return (html_content, attachments).
    Embeds inline images referenced by cid:... as base64 data URIs.
    """
    html_content = None
    text_content = None
    attachments = []

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdisp = str(part.get("Content-Disposition") or "").lower()
            cid = part.get("Content-ID")

            if ctype == "text/html" and "attachment" not in cdisp:
                html_payload = part.get_payload(decode=True)
                html_charset = part.get_content_charset() or "utf-8"
                html_content = html_payload.decode(html_charset, errors="ignore")

            elif ctype == "text/plain" and "attachment" not in cdisp:
                text_payload = part.get_payload(decode=True)
                text_charset = part.get_content_charset() or "utf-8"
                text_content = text_payload.decode(text_charset, errors="ignore")
            else:
                attachments.append(part)
    else:
        ctype = msg.get_content_type()
        if ctype == "text/html":
            html_payload = msg.get_payload(decode=True)
            html_charset = msg.get_content_charset() or "utf-8"
            html_content = html_payload.decode(html_charset, errors="ignore")
        elif ctype == "text/plain":
            text_payload = msg.get_payload(decode=True)
            text_charset = msg.get_content_charset() or "utf-8"
            text_content = text_payload.decode(text_charset, errors="ignore")

    if not html_content:
        if text_content:
            html_content = f"<html><body><pre>{text_content}</pre></body></html>"
        else:
            html_content = "<html><body>(No content)</body></html>"

    # Replace inline images
    for part in attachments:
        cid = part.get("Content-ID")
        if cid:
            cid_clean = cid.strip("<>")
            if cid_clean in html_content:
                payload = part.get_payload(decode=True)
                maintype = part.get_content_maintype()
                subtype = part.get_content_subtype()
                b64 = base64.b64encode(payload).decode("utf-8")
                data_uri = f"data:{maintype}/{subtype};base64,{b64}"

                pattern = re.compile(r'src=["\']cid:' + re.escape(cid_clean) + r'["\']')
                html_content = pattern.sub(f'src="{data_uri}"', html_content)

    return html_content, attachments


def create_attachment_list_html(attachments):
    """Build an HTML snippet listing non-inline attachments at the bottom."""
    attach_names = []
    for part in attachments:
        filename = decode_str(part.get_filename() or "")
        if not filename:
            continue
        cdisp = str(part.get("Content-Disposition") or "").lower()
        cid = part.get("Content-ID")
        # "attachment" in cdisp or no content_id => user-facing attachment
        if "attachment" in cdisp or not cid:
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
    """Use WeasyPrint to convert the HTML to PDF."""
    pdf_bytes = weasyprint.HTML(string=html_content).write_pdf()
    with open(output_pdf_path, 'wb') as f:
        f.write(pdf_bytes)


def append_pdf_attachments_to_pdf(attachments, base_pdf_path):
    """
    Merge PDF attachments into base_pdf_path:
      - Insert a title page for each PDF
      - If merging fails, insert a "broken PDF" page
      - Return (# of PDFs processed, list of broken filenames)
    """
    merger = PyPDF2.PdfMerger()
    merger.append(base_pdf_path)

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

            # Title page
            title_bytes = make_attachment_title_pdf(filename)
            title_path = tmp_dir / f"title_{uuid.uuid4()}.pdf"
            with open(title_path, 'wb') as tf:
                tf.write(title_bytes)
            temp_files.append(title_path)
            merger.append(str(title_path))

            # Actual PDF
            payload = part.get_payload(decode=True)
            pdf_path = tmp_dir / f"attach_{uuid.uuid4()}.pdf"
            with open(pdf_path, 'wb') as pf:
                pf.write(payload)
            temp_files.append(pdf_path)

            try:
                merger.append(str(pdf_path))
            except PyPDF2.errors.PdfReadError as e:
                log(f"[WARNING] Could not merge '{filename}' (PdfReadError): {e}")
                broken_pdfs.append(filename)

                broken_bytes = make_broken_pdf_notice(filename)
                broken_path = tmp_dir / f"broken_{uuid.uuid4()}.pdf"
                with open(broken_path, 'wb') as bf:
                    bf.write(broken_bytes)
                temp_files.append(broken_path)
                merger.append(str(broken_path))

            except Exception as ex:
                log(f"[WARNING] Could not merge '{filename}' (Unknown Error): {ex}")
                broken_pdfs.append(filename)

                broken_bytes = make_broken_pdf_notice(filename)
                broken_path = tmp_dir / f"broken_{uuid.uuid4()}.pdf"
                with open(broken_path, 'wb') as bf:
                    bf.write(broken_bytes)
                temp_files.append(broken_path)
                merger.append(str(broken_path))

    merger.write(base_pdf_path)
    merger.close()

    # Cleanup
    for tmpf in temp_files:
        try:
            os.remove(tmpf)
        except OSError as e:
            log(f"[WARNING] Could not remove temp file {tmpf}: {e}")

    return pdf_count, broken_pdfs


def eml_to_pdf(eml_path, output_pdf_path):
    """
    1) Parse EML -> headers + body + attachments
    2) Insert header block at top, embed inline images
    3) Convert body to PDF
    4) Merge PDF attachments
    5) Return (#pdf_attachments, broken_pdf_list)
    """
    log(f"Reading EML: {eml_path}")
    with open(eml_path, 'rb') as f:
        raw_data = f.read()

    msg = email.message_from_bytes(raw_data, policy=default)

    # Headers
    hdrs = extract_email_headers(msg)

    # Body & attachments
    html_content, attachments = extract_html_and_inline_images(msg)

    # Insert header HTML at top (after <body>)
    header_html = make_header_html(hdrs)
    if "<body>" in html_content:
        html_content = html_content.replace("<body>", f"<body>{header_html}", 1)
    else:
        html_content = f"<html><body>{header_html}{html_content}</body></html>"

    # Insert attachment list near bottom
    html_content += create_attachment_list_html(attachments)

    # Convert to PDF
    log("  -> Converting HTML body to PDF")
    convert_html_to_pdf(html_content, output_pdf_path)

    # Merge PDF attachments
    log("  -> Checking for PDF attachments to merge")
    pdf_count, broken_list = append_pdf_attachments_to_pdf(attachments, output_pdf_path)
    return pdf_count, broken_list


def convert_eml_files_in_directory(eml_dir, output_dir, merge_all=False, report=False):
    """
    Processes .eml files from eml_dir -> PDF in output_dir.
    If merge_all=True, merges them all into merged_all.pdf
    If report=True, writes summary to eml2pdf.txt in output_dir
    """
    entries = [e for e in os.scandir(eml_dir) if e.is_file() and e.name.lower().endswith(".eml")]
    total_files = len(entries)
    if total_files == 0:
        log("No EML files found.")
        return

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    total_pdf_attachments = 0
    broken_pdfs = []
    generated_pdfs = []

    for i, entry in enumerate(entries, start=1):
        log(f"\n=== Processing file {i} of {total_files}: {entry.name} ===")
        out_pdf = os.path.join(output_dir, f"{os.path.splitext(entry.name)[0]}.pdf")
        log(f"Output PDF: {out_pdf}")

        pdf_count, broken_list = eml_to_pdf(entry.path, out_pdf)
        total_pdf_attachments += pdf_count
        broken_pdfs.extend(broken_list)
        generated_pdfs.append(out_pdf)

    merged_filename = None
    if merge_all and generated_pdfs:
        merged_filename = os.path.join(output_dir, "merged_all.pdf")
        log(f"\n[MERGE] Merging all {len(generated_pdfs)} PDFs into {merged_filename}")
        merger = PyPDF2.PdfMerger()
        for pdf_file in generated_pdfs:
            merger.append(pdf_file)
        merger.write(merged_filename)
        merger.close()

    # Summary
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

    for line in summary_lines:
        log(line)

    if report:
        rep_path = os.path.join(output_dir, "eml2pdf.txt")
        with open(rep_path, "w", encoding="utf-8") as repf:
            for line in summary_lines:
                repf.write(line + "\n")
        log(f"[REPORT] Summary written to: {rep_path}")


def main():
    parser = argparse.ArgumentParser(
        description="Convert EML files to PDF (headers, inline images, PDF attachments)."
    )
    parser.add_argument("eml_dir", help="Directory containing .eml files.")
    parser.add_argument("output_dir", help="Directory for generated .pdf files.")
    parser.add_argument("--merge", action="store_true",
                        help="Merge all generated PDFs into 'merged_all.pdf' at the end.")
    parser.add_argument("--report", action="store_true",
                        help="Write a 'eml2pdf.txt' summary report in output_dir.")

    args = parser.parse_args()

    convert_eml_files_in_directory(
        eml_dir=args.eml_dir,
        output_dir=args.output_dir,
        merge_all=args.merge,
        report=args.report
    )


if __name__ == "__main__":
    main()
