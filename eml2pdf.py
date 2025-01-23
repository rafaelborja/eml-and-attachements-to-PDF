#!/usr/bin/env python3
"""
EML to PDF Converter
(C) 2025 Rafael Borja

Features:
- Convert .eml to PDF with inline images.
- Merge PDF attachments with a title page for each attachment.
- Insert a list of attachments at the end of the email body.
- Log progress messages.
- Attempt to silence GLib warnings on Windows.
"""

import os
# Attempt to hide GLib-GIO "Unexpectedly, UWP app..." warnings on Windows
os.environ['G_MESSAGES_DEBUG'] = 'none'

import base64
import email
import re
from email.policy import default
from email.header import decode_header
from pathlib import Path
import argparse
import weasyprint
import PyPDF2


def log(msg: str):
    """Simple logger function."""
    print(f"[INFO] {msg}")


def decode_str(s):
    """Decode MIME-encoded strings, e.g. '=?utf-8?...?='."""
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


def create_attachment_list_html(attachments):
    """
    Generate a small HTML snippet showing all attachments
    (excluding inline images).
    """
    # Filter out inline images (they have content_id and no 'attachment' disposition)
    # We'll call something an "attachment" if:
    #   1) it has a filename, and
    #   2) 'attachment' is in its content disposition, or there's no content_id
    attachment_filenames = []
    for part in attachments:
        filename = decode_str(part.get_filename() or "")
        cd = str(part.get("Content-Disposition") or "").lower()
        content_id = part.get("Content-ID")
        if filename and ("attachment" in cd or not content_id):
            attachment_filenames.append(filename)

    if not attachment_filenames:
        return ""  # No attachments to list

    # Build an HTML list
    items = "".join(f"<li>{fn}</li>" for fn in attachment_filenames)
    html_list = f"""
<hr>
<h3>Attachments in this email:</h3>
<ul>
  {items}
</ul>
"""
    return html_list


def extract_html_and_inline_images(msg):
    """
    Returns:
      - html_content (string) with the message body
      - attachments (list of all parts that might be attachments, including inline images)
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
                # Prefer HTML
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

    if not html_content:
        if text_content:
            html_content = f"<html><body><pre>{text_content}</pre></body></html>"
        else:
            html_content = "<html><body>(No content)</body></html>"

    # Embed inline images
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


def convert_html_to_pdf(html_content, output_pdf_path):
    """Convert the given HTML string to a PDF using WeasyPrint."""
    pdf_bytes = weasyprint.HTML(string=html_content).write_pdf()
    with open(output_pdf_path, 'wb') as f:
        f.write(pdf_bytes)


def make_attachment_title_pdf(attachment_name: str) -> bytes:
    """
    Create a single-page PDF (in memory) that has a header "Attachment: <filename>".
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


def append_pdf_attachments_to_pdf(attachments, base_pdf_path):
    """
    Merges PDF attachments into base_pdf_path.
    Before each PDF, insert a single-page PDF with the filename at the top.
    """
    merger = PyPDF2.PdfMerger()
    merger.append(base_pdf_path)

    # Keep track of temporary files to delete later
    temp_files = []

    for part in attachments:
        filename = decode_str(part.get_filename() or "")
        # Decide whether it's actually a PDF attachment
        if filename.lower().endswith(".pdf"):
            log(f"  -> Processing PDF attachment: {filename}")

            # 1) Make a single page "attachment name" PDF
            attachment_title_pdf = make_attachment_title_pdf(filename)
            title_pdf_temp = base_pdf_path + ".title_tmp.pdf"
            with open(title_pdf_temp, 'wb') as tf:
                tf.write(attachment_title_pdf)
            temp_files.append(title_pdf_temp)

            # Merge the single-page title PDF
            merger.append(title_pdf_temp)

            # 2) Write the actual PDF attachment to a temp file
            payload = part.get_payload(decode=True)
            pdf_temp_path = base_pdf_path + ".attach_tmp.pdf"
            with open(pdf_temp_path, 'wb') as f:
                f.write(payload)
            temp_files.append(pdf_temp_path)

            try:
                # Merge the actual attachment PDF
                merger.append(pdf_temp_path)
            except PyPDF2.errors.PdfReadError as e:
                log(f"[WARNING] Unreadable PDF '{filename}': {e}")
                broken_pdf_bytes = make_broken_pdf_notice(filename)
                # Write temp file
                broken_temp_path = base_pdf_path + ".broken_notice_tmp.pdf"
                with open(broken_temp_path, 'wb') as bf:
                    bf.write(broken_pdf_bytes)
                merger.append(broken_temp_path)
                temp_files.append(broken_temp_path)
                continue

    # Write out the final merged PDF and close
    merger.write(base_pdf_path)
    merger.close()

    # Remove temp files now that the merger is done with them
    for tmpf in temp_files:
        try:
            os.remove(tmpf)
        except OSError as e:
            print(f"[WARNING] Could not remove temp file {tmpf}: {e}")

def make_broken_pdf_notice(filename):
    html = f"""
    <html>
      <body>
        <h1 style="color:red;">Could not merge attachment: {filename}</h1>
        <p>File might be corrupted or unsupported.</p>
      </body>
    </html>
    """
    return weasyprint.HTML(string=html).write_pdf()

def eml_to_pdf(eml_path, output_pdf_path):
    """
    Main routine:
      1. Parse the .eml file
      2. Extract HTML/text + inline images
      3. Append a list of attachments at the end of the HTML
      4. Convert to a base PDF
      5. Insert "title pages" for each PDF attachment, then merge attachments
    """
    log(f"Reading EML: {eml_path}")
    with open(eml_path, 'rb') as f:
        raw_data = f.read()

    msg = email.message_from_bytes(raw_data, policy=default)
    html_content, attachments = extract_html_and_inline_images(msg)

    # Add attachment list to the end of the email content
    attachment_list_html = create_attachment_list_html(attachments)
    if attachment_list_html:
        html_content += attachment_list_html

    # Convert the email body to a PDF
    log("  -> Converting HTML body to PDF")
    convert_html_to_pdf(html_content, output_pdf_path)

    # Merge PDF attachments
    log("  -> Checking for PDF attachments to merge")
    append_pdf_attachments_to_pdf(attachments, output_pdf_path)


def convert_eml_files_in_directory(eml_dir, output_dir):
    """
    Processes all .eml files in `eml_dir` and writes .pdf files to `output_dir`.
    """
    entries = [e for e in os.scandir(eml_dir) if e.is_file() and e.name.lower().endswith(".eml")]
    total = len(entries)
    if total == 0:
        log("No EML files found.")
        return

    Path(output_dir).mkdir(parents=True, exist_ok=True)

    for i, entry in enumerate(entries):
        log(f"\n=== Processing file {i+1} of {total}: {entry.name} ===")
        base_name = os.path.splitext(entry.name)[0]
        output_pdf_path = os.path.join(output_dir, f"{base_name}.pdf")
        log(f"Output PDF: {output_pdf_path}")
        eml_to_pdf(entry.path, output_pdf_path)


def main():
    parser = argparse.ArgumentParser(
        description="Convert EML files to PDF (inline images + PDF attachments)"
    )
    parser.add_argument("eml_dir", help="Path to the directory containing .eml files.")
    parser.add_argument("output_dir", help="Directory where .pdf files will be saved.")
    args = parser.parse_args()

    convert_eml_files_in_directory(args.eml_dir, args.output_dir)


if __name__ == "__main__":
    main()
