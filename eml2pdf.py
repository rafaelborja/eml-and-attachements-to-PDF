import os
import base64
import email
import re
import uuid
import argparse
import logging
import hashlib
from concurrent.futures import ThreadPoolExecutor, as_completed
from email.policy import default
from email.header import decode_header
from email.utils import getaddresses, parsedate_to_datetime
from pathlib import Path
import mailbox

import PyPDF2
import weasyprint

try:
    import html2text
except ImportError:
    html2text = None

# --- Logger Setup ---
logger = logging.getLogger(__name__)
DEBUG_MODE = False

# --- Helper Functions ---

def get_sha256_hash(content_bytes: bytes) -> str:
    """Computes the SHA256 hash of byte content."""
    sha256 = hashlib.sha256()
    sha256.update(content_bytes)
    return sha256.hexdigest()

def get_pdf_size_mb(pdf_path):
    try:
        file_size_bytes = os.path.getsize(pdf_path)
        return file_size_bytes / (1024 * 1024)
    except OSError as e:
        logger.warning(f"Could not get size for {pdf_path}: {e}", exc_info=DEBUG_MODE)
        return 0

def get_pdf_page_count(pdf_path):
    try:
        with open(pdf_path, 'rb') as pdf_file:
            pdf_reader = PyPDF2.PdfReader(pdf_file)
            return len(pdf_reader.pages)
    except FileNotFoundError:
        logger.warning(f"File not found when getting page count: {pdf_path}")
        return 0
    except PyPDF2.errors.PdfReadError as e: # More specific PyPDF2 error
        logger.warning(f"PyPDF2 error reading {pdf_path} for page count: {e}", exc_info=DEBUG_MODE)
        return 0
    except Exception as e: # General fallback
        logger.warning(f"Could not get page count for {pdf_path}: {e}", exc_info=DEBUG_MODE)
        return 0

def extract_email_headers(msg) -> dict:
    def decode_str(s):
        if not s: return ""
        parts = decode_header(s)
        decoded = ""
        for text, enc in parts:
            if enc:
                try: decoded += text.decode(enc, errors='replace')
                except LookupError: decoded += text.decode('utf-8', errors='replace')
            elif isinstance(text, bytes): decoded += text.decode('utf-8', errors='replace')
            else: decoded += text
        return decoded

    def parse_address_header(value: str) -> str:
        if not value: return ""
        pairs = getaddresses([value])
        # Ensure names are decoded and handle cases where name might be None
        result_list = []
        for raw_name, email_addr in pairs:
            name_decoded = decode_str(raw_name)
            if name_decoded and email_addr:
                result_list.append(f"{name_decoded} <{email_addr}>")
            elif email_addr: # Only email address
                result_list.append(email_addr)
            elif name_decoded: # Only name (less common for From/To but possible)
                result_list.append(name_decoded)
        return ", ".join(filter(None, result_list))


    return {
        "date": decode_str(msg.get("Date", "")),
        "from": parse_address_header(msg.get("From", "")),
        "to": parse_address_header(msg.get("To", "")),
        "cc": parse_address_header(msg.get("Cc", "")),
        "subject": decode_str(msg.get("Subject", ""))
    }

def make_header_html(headers: dict) -> str:
    header_style = "font-family:sans-serif; border-bottom:1px solid #ccc; padding:10px; margin-bottom:10px; background-color:#f9f9f9;"
    p_style = "font-size:small; margin:4px 0; word-wrap:break-word;" # Ensure long lines wrap
    strong_style = "color:#333;"
    html_parts = [f"<div style='{header_style}'>"]
    for key, label in [('date', 'Date'), ('from', 'From'), ('to', 'To'), ('cc', 'Cc'), ('subject', 'Subject')]:
        if headers.get(key): # Check if header exists
            # Escape HTML content from headers to prevent XSS or rendering issues
            escaped_value = html2text.html.escape(headers[key])
            html_parts.append(f"<p style='{p_style}'><strong style='{strong_style}'>{label}:</strong> {escaped_value}</p>")
    html_parts.append("</div>")
    return "\n".join(html_parts)

def create_attachment_name_list_html(attachments_data) -> str:
    """Creates an HTML list of attachment names and their SHA256 hashes from structured attachment data."""
    if not attachments_data:
        return ""
    items = "".join(f"<li>{html2text.html.escape(att['filename'])} (SHA256: {att['hash']})</li>" for att in attachments_data)
    return f"<hr><h3>Attachments Processed (Summary):</h3><ul>{items}</ul>"


def convert_html_to_pdf(html_content, output_pdf_path, base_url_for_assets=None):
    """Converts HTML string to a PDF file."""
    try:
        path_str = str(output_pdf_path) # Ensure it's a string
        # Providing a base_url helps WeasyPrint resolve relative paths for images, CSS if any were in HTML
        # If HTML is self-contained (e.g., with base64 images), base_url is less critical but good practice.
        # Defaulting to the current directory if not specified.
        effective_base_url = base_url_for_assets if base_url_for_assets else Path(".").resolve().as_uri() + "/"
        html = weasyprint.HTML(string=html_content, base_url=effective_base_url)
        html.write_pdf(path_str)
        logger.debug(f"Successfully converted HTML to PDF: {path_str}")
    except Exception as e:
        logger.error(f"PDF conversion failed for {output_pdf_path}: {e}", exc_info=DEBUG_MODE)
        with open(str(output_pdf_path), 'wb') as f: # Ensure path_str is used here too
            f.write(b"%PDF-1.0\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj 2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj 3 0 obj<</Type/Page/MediaBox[0 0 612 792]>>endobj\nxref\n0 4\n0000000000 65535 f \n0000000009 00000 n \n0000000052 00000 n \n0000000101 00000 n \ntrailer<</Size 4/Root 1 0 R>>\nstartxref\n140\n%%EOF")


def extract_email_main_html_and_attachments(msg):
    html_body = None
    text_body = None
    attachments_data_list = []
    inline_images = {}

    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdisp = str(part.get("Content-Disposition") or "").lower()
            cid = part.get("Content-ID", "").strip("<>") # Ensure CID is clean
            filename = part.get_filename()

            is_attachment_file = "attachment" in cdisp or (filename and not cid and part.get_content_maintype() != 'multipart')

            if cid and part.get_content_maintype() == 'image' and ("inline" in cdisp or not cdisp and not filename):
                try:
                    image_data = part.get_payload(decode=True)
                    image_type = ctype.split('/')[-1]
                    if image_type and image_data: # Ensure data and type are valid
                        base64_data = base64.b64encode(image_data).decode()
                        inline_images[cid] = f"data:image/{image_type};base64,{base64_data}"
                except Exception as e:
                    logger.warning(f"Could not embed inline image {cid}: {e}", exc_info=DEBUG_MODE)
                continue # Skip further processing for this inline part

            if ctype == "text/html" and not is_attachment_file and not html_body : # Prioritize first HTML part
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8" # Default to utf-8
                try: html_body = payload.decode(charset, errors="replace")
                except LookupError: html_body = payload.decode("utf-8", errors="replace") # Fallback
            elif ctype == "text/plain" and not is_attachment_file and not text_body: # Prioritize first text part
                payload = part.get_payload(decode=True)
                charset = part.get_content_charset() or "utf-8"
                try: text_body = payload.decode(charset, errors="replace")
                except LookupError: text_body = payload.decode("utf-8", errors="replace")
            # Check if it's a file to be processed as an attachment (has filename and actual content)
            elif filename and part.get_payload(decode=False) is not None: # decode=False to check existence
                # Ensure it's not a multipart alternative or related part unless it's explicitly an attachment
                if "attachment" in cdisp or (not cdisp and part.get_content_maintype() != 'multipart'):
                    try:
                        content_bytes = part.get_payload(decode=True)
                        if content_bytes: # Ensure there's actual content
                            sha256_hash = get_sha256_hash(content_bytes)
                            attachments_data_list.append({
                                'filename': filename,
                                'content': content_bytes,
                                'hash': sha256_hash
                            })
                            logger.debug(f"Extracted attachment: {filename}, SHA256: {sha256_hash[:8]}...")
                    except Exception as e:
                        logger.warning(f"Failed to process attachment part '{filename}': {e}", exc_info=DEBUG_MODE)
    else: # Non-multipart message
        payload = msg.get_payload(decode=True)
        charset = msg.get_content_charset() or "utf-8"
        ctype = msg.get_content_type()
        if ctype == "text/html":
            try: html_body = payload.decode(charset, errors="replace")
            except LookupError: html_body = payload.decode("utf-8", errors="replace")
        elif ctype == "text/plain":
            try: text_body = payload.decode(charset, errors="replace")
            except LookupError: text_body = payload.decode("utf-8", errors="replace")
        # If it's a non-multipart message but has a filename, treat the payload as an attachment.
        elif msg.get_filename():
            content_bytes = payload
            if content_bytes: # Ensure there's actual content
                sha256_hash = get_sha256_hash(content_bytes)
                attachments_data_list.append({
                    'filename': msg.get_filename(),
                    'content': content_bytes,
                    'hash': sha256_hash
                })
                logger.debug(f"Extracted attachment from non-multipart: {msg.get_filename()}, SHA256: {sha256_hash[:8]}...")


    final_html_content = ""
    if html_body:
        final_html_content = html_body
        # Embed inline images using CIDs
        for cid_val, base64_src in inline_images.items():
            # Replace both cid:cid_val and "cid_val" (sometimes found in src attributes without cid: prefix)
            final_html_content = final_html_content.replace(f"cid:{cid_val}", base64_src, 100)
            final_html_content = final_html_content.replace(f'"{cid_val}"', f'"{base64_src}"', 100) # For src="cid_val" cases
    elif text_body:
        escaped_text = html2text.html.escape(text_body) # Escape plain text for HTML <pre>
        final_html_content = f"<html><head><meta charset='UTF-8'></head><body><pre>{escaped_text}</pre></body></html>"
    else: # No body content found
        final_html_content = "<html><head><meta charset='UTF-8'></head><body><p>(No textual content found in email body)</p></body></html>"

    return final_html_content, attachments_data_list


def append_attachment_pages_to_merger(merger: PyPDF2.PdfMerger,
                                      attachments_data_list: list,
                                      seen_hashes_in_document: dict,
                                      deduplicate_active: bool,
                                      tmp_processing_dir: Path):
    attachments_processed_count = 0
    broken_attachment_list = []
    reused_attachments_this_call = 0 # New counter
    temp_files_created = []

    for attachment_data in attachments_data_list:
        current_filename = attachment_data['filename']
        current_hash = attachment_data['hash']
        current_content = attachment_data['content']
        attachments_processed_count += 1

        header_page_html_parts = [
            "<html><head><meta charset='UTF-8'>",
            "<style>",
            "body { margin: 0; display: flex; flex-direction: column; justify-content: center; align-items: center; height: 100vh; text-align: center; font-family: sans-serif; border: 2px dashed #ccc; box-sizing: border-box; padding: 20px; }",
            ".container { word-wrap: break-word; max-width: 90%; }",
            "h1 { font-size: 16pt; color: #333; margin-bottom: 10px; }",
            "p { font-size: 10pt; color: #555; margin-bottom: 5px; }",
            ".sha { font-family: monospace; font-size: 9pt; word-break: break-all; }", # Ensure SHA hash wraps
            ".note { font-weight: bold; color: #d9534f; margin-top: 15px; font-size: 11pt; }",
            "</style></head><body><div class='container'>",
            f"<h1>Attachment File:</h1><p>{html2text.html.escape(current_filename)}</p>", # Escape filename
            f"<h2>SHA256 Content ID:</h2><p class='sha'>{current_hash}</p>"
        ]
        is_duplicate_ref = False

        if deduplicate_active and current_hash in seen_hashes_in_document:
            first_occurrence_info = seen_hashes_in_document[current_hash]
            first_occurrence_filename = first_occurrence_info['filename']
            duplicate_message = (
                f"DUPLICATE CONTENT: This attachment ('{html2text.html.escape(current_filename)}') "
                f"has the exact same content as attachment '{html2text.html.escape(first_occurrence_filename)}' "
                f"(ID: {current_hash}) previously in this document part. "
                "The content is not repeated here."
            )
            header_page_html_parts.append(f"<p class='note'>{duplicate_message}</p>")
            is_duplicate_ref = True
            reused_attachments_this_call += 1 # Increment reused counter

        header_page_html_parts.extend(["</div></body></html>"])
        header_html_full = "\n".join(header_page_html_parts)

        header_pdf_path = tmp_processing_dir / f"header_{uuid.uuid4()}.pdf"
        temp_files_created.append(header_pdf_path)
        try:
            convert_html_to_pdf(header_html_full, header_pdf_path)
            if header_pdf_path.exists() and header_pdf_path.stat().st_size > 0:
                merger.append(str(header_pdf_path))
            else:
                logger.warning(f"Header PDF for {current_filename} is empty or not created.")
                broken_attachment_list.append(f"{current_filename} (header page error)")
        except Exception as e_header:
            logger.warning(f"Could not create/append info PDF for {current_filename}: {e_header}", exc_info=DEBUG_MODE)
            broken_attachment_list.append(f"{current_filename} (header page creation error: {e_header})")
            continue

        if not is_duplicate_ref:
            if current_filename.lower().endswith(".pdf") and current_content:
                actual_attachment_pdf_path = tmp_processing_dir / f"attach_content_{uuid.uuid4()}.pdf"
                temp_files_created.append(actual_attachment_pdf_path)
                try:
                    with open(actual_attachment_pdf_path, 'wb') as f_attach: f_attach.write(current_content)
                    try: # Validate before appending
                        PyPDF2.PdfReader(str(actual_attachment_pdf_path))
                        merger.append(str(actual_attachment_pdf_path))
                    except PyPDF2.errors.PdfReadError:
                        logger.warning(f"Attachment '{current_filename}' is a broken PDF. Creating error placeholder.")
                        broken_attachment_list.append(f"{current_filename} (corrupted PDF content)")
                        error_page_html = f"<html><body><p style='color:red;'>Error: Attachment '{html2text.html.escape(current_filename)}' (SHA256: {current_hash}) could not be embedded as it appears to be a corrupted PDF.</p></body></html>"
                        error_pdf_path = tmp_processing_dir / f"error_attach_{uuid.uuid4()}.pdf"
                        temp_files_created.append(error_pdf_path)
                        convert_html_to_pdf(error_page_html, error_pdf_path)
                        if error_pdf_path.exists() and error_pdf_path.stat().st_size > 0: merger.append(str(error_pdf_path))
                except Exception as e_attach_content:
                    logger.warning(f"Could not save/append PDF attachment content {current_filename}: {e_attach_content}", exc_info=DEBUG_MODE)
                    broken_attachment_list.append(f"{current_filename} (content append error: {e_attach_content})")

            if deduplicate_active: # Record only if content was actually embedded
                seen_hashes_in_document[current_hash] = {'filename': current_filename}
        else:
            logger.info(f"Handled attachment '{current_filename}' (SHA256: {current_hash}) as a duplicate reference.")

    for tmpf in temp_files_created:
        try:
            if tmpf.exists(): os.remove(tmpf)
        except OSError: logger.warning(f"Could not remove temporary file: {tmpf}", exc_info=DEBUG_MODE)

    return attachments_processed_count, broken_attachment_list, reused_attachments_this_call


def eml_to_output_for_individual_file(msg, output_path, out_format, deduplicate_attachments_active, tmp_dir_base):
    headers = extract_email_headers(msg)
    email_body_html, attachments_data = extract_email_main_html_and_attachments(msg)

    full_html_content = make_header_html(headers) + email_body_html
    full_html_content += create_attachment_name_list_html(attachments_data)

    reused_attach_count_total = 0

    if out_format == "pdf":
        current_processing_tmp_dir = Path(tmp_dir_base) / f"indiv_{uuid.uuid4()}"
        current_processing_tmp_dir.mkdir(parents=True, exist_ok=True)
        temp_files_to_clean_main = []

        body_render_html = full_html_content + "<hr style='page-break-after: always; visibility: hidden;'>"
        email_body_only_pdf_path = current_processing_tmp_dir / "main_email_body.pdf"
        temp_files_to_clean_main.append(email_body_only_pdf_path)
        convert_html_to_pdf(body_render_html, email_body_only_pdf_path, base_url_for_assets=current_processing_tmp_dir.as_uri())


        final_merger = PyPDF2.PdfMerger()
        if email_body_only_pdf_path.exists() and email_body_only_pdf_path.stat().st_size > 0:
            final_merger.append(str(email_body_only_pdf_path))
        else:
            logger.warning(f"Main email body PDF for {output_path} is missing or empty.")

        seen_hashes_this_document = {}
        pdf_attachments_data = [att for att in attachments_data if att['filename'].lower().endswith(".pdf")]

        # shared_tmp_weasy_artifacts_dir is passed as tmp_processing_dir to append_attachment_pages_to_merger
        # as it's where header PDFs etc. will be made.
        att_count, broken_atts, reused_attach_count_total = append_attachment_pages_to_merger(
            final_merger, pdf_attachments_data, seen_hashes_this_document,
            deduplicate_attachments_active, Path("tmp_eml_attachments")
        )

        try:
            final_merger.write(str(output_path))
        except Exception as e:
            logger.error(f"Failed to write final PDF {output_path}: {e}", exc_info=DEBUG_MODE)
        finally:
            final_merger.close()

        for tmpf in temp_files_to_clean_main:
            if tmpf.exists():
                try: os.remove(tmpf)
                except OSError as e: logger.warning(f"Could not remove temp file {tmpf}: {e}")
        try: # Cleanup specific temp dir if empty
            if current_processing_tmp_dir.exists() and not any(current_processing_tmp_dir.iterdir()):
                current_processing_tmp_dir.rmdir()
        except OSError as e: logger.warning(f"Could not remove temp dir {current_processing_tmp_dir}: {e}")

        return att_count, broken_atts, reused_attach_count_total # Return reused count

    # ... (HTML/MD output, unchanged from previous version, returning 0 for reused_attach_count_total)
    elif out_format == "html":
        try:
            with open(output_path, "w", encoding="utf-8") as f: f.write(full_html_content)
        except Exception as e: logger.error(f"Failed to write HTML {output_path}: {e}", exc_info=DEBUG_MODE)
        return 0, [], 0
    elif out_format == "md":
        if not html2text: raise RuntimeError("html2text not installed.")
        try:
            converter = html2text.HTML2Text(bodywidth=0); md_output = converter.handle(full_html_content)
            with open(output_path, "w", encoding="utf-8") as f: f.write(md_output)
        except Exception as e: logger.error(f"Failed to write MD {output_path}: {e}", exc_info=DEBUG_MODE)
        return 0, [], 0
    return 0, [], 0


def process_email_task(task_info, output_dir_path, out_format, prepend_date,
                       deduplicate_attachments_active, merge_all_pdf_mode,
                       task_temp_dir_base):
    msg_obj = None
    base_name_for_file = ""
    identifier_for_log = ""
    attachments_data = []

    try:
        if task_info['type'] == 'eml':
            # ... (EML loading logic) ...
            entry_path = Path(task_info['path'])
            identifier_for_log = entry_path.name
            base_name_for_file = entry_path.stem
            with open(entry_path, 'rb') as f: raw_data = f.read()
            msg_obj = email.message_from_bytes(raw_data, policy=default)
        elif task_info['type'] == 'mbox':
            # ... (MBOX message retrieval) ...
            msg_obj = task_info['msg_obj']
            mbox_filename = task_info['mbox_filename']
            msg_index = task_info['msg_index']
            identifier_for_log = f"message {msg_index} from {Path(mbox_filename).name}"
            base_name_for_file = f"{Path(mbox_filename).stem}_message_{msg_index}"
        else: # Should not happen
            logger.error(f"Unknown task type: {task_info.get('type')}")
            # Adjust return tuple arity based on mode
            return (None, [], identifier_for_log) if merge_all_pdf_mode else (None, 0, [], 0, identifier_for_log)


        if not msg_obj:
            logger.warning(f"No message object for {identifier_for_log}.")
            return (None, [], identifier_for_log) if merge_all_pdf_mode else (None, 0, [], 0, identifier_for_log)

        headers = extract_email_headers(msg_obj)
        email_body_html, attachments_data = extract_email_main_html_and_attachments(msg_obj)

        date_prefix_str = ""
        if prepend_date:
            date_val = headers.get("date", "")
            if date_val:
                try:
                    dt_parsed = parsedate_to_datetime(date_val)
                    if dt_parsed: date_prefix_str = dt_parsed.strftime("%Y-%m-%d") + "_"
                except Exception: logger.warning(f"Could not parse date '{date_val}' for {identifier_for_log}", exc_info=DEBUG_MODE)

        safe_base_name = re.sub(r'[\\/*?:"<>|]', "_", base_name_for_file)

        if merge_all_pdf_mode and out_format == 'pdf':
            task_specific_tmp_dir = Path(task_temp_dir_base) / f"task_{uuid.uuid4()}"
            task_specific_tmp_dir.mkdir(parents=True, exist_ok=True)
            body_render_html = make_header_html(headers) + email_body_html + create_attachment_name_list_html(attachments_data)
            body_render_html += "<hr style='page-break-after: always; visibility: hidden;'>"
            body_only_pdf_filename = f"{date_prefix_str}{safe_base_name}_body.pdf"
            body_only_pdf_path = task_specific_tmp_dir / body_only_pdf_filename
            convert_html_to_pdf(body_render_html, body_only_pdf_path, base_url_for_assets=task_specific_tmp_dir.as_uri())
            if not body_only_pdf_path.exists() or body_only_pdf_path.stat().st_size == 0:
                logger.error(f"Body PDF generation failed for {identifier_for_log} at {body_only_pdf_path}")
                return None, attachments_data, identifier_for_log
            return str(body_only_pdf_path), attachments_data, identifier_for_log
        else:
            ext = {"pdf": ".pdf", "html": ".html", "md": ".md"}.get(out_format, ".pdf")
            out_filename = f"{date_prefix_str}{safe_base_name}{ext}"
            final_output_path = output_dir_path / out_filename

            # Pass tmp_dir_base for individual file processing to create its own sub-temp dir
            pdf_attach_count, broken_list, reused_count = eml_to_output_for_individual_file(
                msg_obj, final_output_path, out_format,
                deduplicate_attachments_active,
                task_temp_dir_base
            )
            return str(final_output_path), pdf_attach_count, broken_list, reused_count, identifier_for_log

    except Exception as e:
        logger.error(f"Unhandled error in task {identifier_for_log}: {e}", exc_info=DEBUG_MODE)
        return (None, [], identifier_for_log) if merge_all_pdf_mode else (None, 0, [], 0, identifier_for_log)


def convert_eml_files_in_directory(eml_dir, output_dir,
                                   merge_all=False, report=False,
                                   prepend_date=False,
                                   out_format="pdf",
                                   mbox_mode=False,
                                   splitsize=0,
                                   splitpages=0,
                                   num_threads=8,
                                   deduplicate_attachments=True):

    output_path_obj = Path(output_dir)
    output_path_obj.mkdir(parents=True, exist_ok=True)

    shared_tmp_weasy_artifacts_dir = Path("tmp_eml_attachments") # For headers, small PDFs by WeasyPrint
    shared_tmp_weasy_artifacts_dir.mkdir(parents=True, exist_ok=True)
    tasks_temp_dir_base = Path("tmp_task_processing") # For larger intermediate files like body PDFs
    tasks_temp_dir_base.mkdir(parents=True, exist_ok=True)

    tasks_to_submit = []
    input_items_count = 0
    # ... (EML/MBOX scanning logic as before) ...
    if not mbox_mode:
        try:
            entries = [e for e in os.scandir(eml_dir) if e.is_file() and e.name.lower().endswith(".eml")]
            input_items_count = len(entries)
            if input_items_count == 0: logger.info("No EML files found.") # (handle empty report)
            for entry in entries: tasks_to_submit.append({'type': 'eml', 'path': str(entry.path)})
        except Exception as e: logger.error(f"Error scanning EML dir {eml_dir}: {e}", exc_info=DEBUG_MODE); return
    else: # MBOX mode
        try:
            mbox_entries = [e for e in os.scandir(eml_dir) if e.is_file() and e.name.lower().endswith((".mbox", ".mbx"))]
            if not mbox_entries: logger.info("No MBOX files found.") # (handle empty report)
            for mbox_file_entry in mbox_entries:
                mbox_path = Path(mbox_file_entry.path)
                try:
                    mbox = mailbox.mbox(str(mbox_path), factory=None, create=False)
                    num_messages_in_mbox = 0
                    for i, key in enumerate(list(mbox.iterkeys())):
                        try:
                            msg = mbox.get_message(key)
                            tasks_to_submit.append({'type': 'mbox', 'msg_obj': msg, 'mbox_filename': mbox_file_entry.name, 'msg_index': i + 1})
                            num_messages_in_mbox +=1
                        except Exception as e_msg: logger.error(f"Failed to get/parse message key {key} from {mbox_file_entry.name}: {e_msg}", exc_info=DEBUG_MODE)
                    input_items_count += num_messages_in_mbox
                    mbox.close()
                except Exception as e_mbox: logger.error(f"Error processing MBOX file {mbox_path}: {e_mbox}", exc_info=DEBUG_MODE)
        except Exception as e: logger.error(f"Error scanning MBOX dir {eml_dir}: {e}", exc_info=DEBUG_MODE); return

    if not tasks_to_submit: logger.info("No email items to process."); return
    logger.info(f"Found {input_items_count} total items to process using up to {num_threads} threads.")

    prepare_for_final_pdf_merge = (merge_all and out_format == 'pdf')
    processed_task_results = []
    with ThreadPoolExecutor(max_workers=num_threads, thread_name_prefix='EMLWorker') as executor:
        futures = { executor.submit(process_email_task, task, output_path_obj, out_format, prepend_date, deduplicate_attachments, prepare_for_final_pdf_merge, tasks_temp_dir_base): task for task in tasks_to_submit }
        for i, future in enumerate(as_completed(futures)):
            # ... (error handling for future.result as before) ...
            task_info_orig = futures[future]
            log_id = task_info_orig.get('path', task_info_orig.get('mbox_filename', 'N/A')) # For logging
            try:
                result = future.result()
                processed_task_results.append(result)
            except Exception as e: logger.error(f"Task for '{log_id}' generated an unhandled exception: {e}", exc_info=DEBUG_MODE)
            logger.info(f"Progress: {i+1}/{len(tasks_to_submit)} tasks processed.")


    total_pdf_attachments_mentioned = 0
    all_broken_pdf_attachments = []
    final_generated_files_for_report = []
    temp_body_pdfs_to_clean = []
    total_reused_attachments_count = 0 # Initialize counter for reused attachments

    if prepare_for_final_pdf_merge:
        processed_task_results.sort(key=lambda x: x[2]) # Sort by original_id
        current_batch_items_for_merger = []
        current_accumulated_size_mb = 0
        current_accumulated_page_count = 0
        split_file_counter = 1

        def _execute_final_merge_batch(counter_val):
            nonlocal current_batch_items_for_merger, current_accumulated_size_mb, current_accumulated_page_count, final_generated_files_for_report, total_pdf_attachments_mentioned, all_broken_pdf_attachments, temp_body_pdfs_to_clean, total_reused_attachments_count # Add total_reused_attachments_count
            if not current_batch_items_for_merger: return
            output_merge_name = f"merged_output_part_{counter_val:03d}.pdf"
            output_merge_path = output_path_obj / output_merge_name
            batch_merger = PyPDF2.PdfMerger()
            seen_hashes_this_merged_part = {}

            for body_pdf_path_str, attachments_data_list_for_email in current_batch_items_for_merger:
                if body_pdf_path_str:
                    body_pdf_path = Path(body_pdf_path_str)
                    if body_pdf_path.exists() and body_pdf_path.stat().st_size > 0:
                        batch_merger.append(str(body_pdf_path))
                        temp_body_pdfs_to_clean.append(body_pdf_path)

                pdf_attachments_for_this_email = [att for att in attachments_data_list_for_email if att['filename'].lower().endswith(".pdf")]
                if pdf_attachments_for_this_email:
                    att_count, broken_atts, reused_count_batch = append_attachment_pages_to_merger( # Get reused_count
                        batch_merger, pdf_attachments_for_this_email, seen_hashes_this_merged_part,
                        deduplicate_attachments, shared_tmp_weasy_artifacts_dir
                    )
                    total_pdf_attachments_mentioned += att_count
                    all_broken_pdf_attachments.extend(broken_atts)
                    total_reused_attachments_count += reused_count_batch # Accumulate reused count

            try:
                batch_merger.write(str(output_merge_path))
                final_generated_files_for_report.append(str(output_merge_path))
            except Exception as e: logger.error(f"Failed to write merged part {output_merge_path}: {e}", exc_info=DEBUG_MODE)
            finally: batch_merger.close()
            current_batch_items_for_merger, current_accumulated_size_mb, current_accumulated_page_count = [], 0, 0

        for body_pdf_path, attachments_data, _original_id in processed_task_results:
            if not body_pdf_path: continue
            body_pdf_size_mb = get_pdf_size_mb(body_pdf_path)
            body_pdf_page_count = get_pdf_page_count(body_pdf_path)
            estimated_attach_pages = sum(1 for att in attachments_data if att['filename'].lower().endswith(".pdf"))
            if current_batch_items_for_merger:
                if splitsize > 0 and current_accumulated_size_mb + body_pdf_size_mb > splitsize:
                    _execute_final_merge_batch(split_file_counter); split_file_counter += 1
                elif splitpages > 0 and current_accumulated_page_count + body_pdf_page_count + estimated_attach_pages > splitpages:
                    _execute_final_merge_batch(split_file_counter); split_file_counter += 1
            current_batch_items_for_merger.append((body_pdf_path, attachments_data))
            current_accumulated_size_mb += body_pdf_size_mb
            current_accumulated_page_count += (body_pdf_page_count + estimated_attach_pages)
        _execute_final_merge_batch(split_file_counter)
    else: # Individual outputs
        # Result: (final_output_path, pdf_attach_count, broken_list, reused_attach_count, original_id)
        for res in processed_task_results:
            if res and res[0]: # Valid output path
                final_generated_files_for_report.append(res[0])
                total_pdf_attachments_mentioned += res[1]
                all_broken_pdf_attachments.extend(res[2])
                total_reused_attachments_count += res[3] # Add reused count

    # --- Summary Report ---
    successful_conversions_count = sum(1 for f in final_generated_files_for_report if Path(f).exists() and Path(f).stat().st_size > 0)
    items_attempted = input_items_count
    failed_conversions_count = items_attempted - successful_conversions_count

    summary_lines = [
        "\n=== CONVERSION SUMMARY REPORT ===",
        # ... (other summary lines as before) ...
        f"Attachment Deduplication Active: {'Yes' if deduplicate_attachments else 'No'}"
    ]
    if out_format == "pdf":
        summary_lines.append(f"Total PDF attachments processed (mentions/references): {total_pdf_attachments_mentioned}")
        if deduplicate_attachments: # Only show reused count if feature was active
            summary_lines.append(f"PDF Attachments Reused (deduplicated references): {total_reused_attachments_count}")
        if all_broken_pdf_attachments:
            summary_lines.append(f"Broken/Unreadable PDF attachment contents encountered ({len(all_broken_pdf_attachments)}):")
            summary_lines.extend([f"  - {html2text.html.escape(fn)}" for fn in all_broken_pdf_attachments])
        else: summary_lines.append("No broken PDF attachment contents encountered.")
    # ... (rest of summary lines for merge, splitsize etc.) ...
    if merge_all and out_format == 'pdf':
        summary_lines.append(f"Output PDF(s) merged into {len(final_generated_files_for_report)} file(s):")
        for fpath in final_generated_files_for_report: summary_lines.append(f"  - {Path(fpath).name}")
        if splitsize > 0: summary_lines.append(f"  (Split by size: {splitsize} MB)")
        if splitpages > 0: summary_lines.append(f"  (Split by pages: {splitpages})")
    else:
        summary_lines.append(f"Individual output files generated: {len(final_generated_files_for_report)}")

    for line in summary_lines: logger.info(line)
    if report:
        rep_path = output_path_obj / "eml_conversion_report.txt"
        try:
            with open(rep_path, "w", encoding="utf-8") as repf: repf.write("\n".join(summary_lines))
            logger.info(f"Summary report written to: {rep_path.resolve()}")
        except Exception as e: logger.error(f"Failed to write summary report to {rep_path}: {e}", exc_info=DEBUG_MODE)

    # ... (Cleanup logic for temp_body_pdfs_to_clean and tasks_temp_dir_base as before) ...
    for f_path in temp_body_pdfs_to_clean:
        try:
            if f_path.exists(): os.remove(f_path)
        except OSError as e: logger.warning(f"Could not remove temp body PDF {f_path}: {e}")
    try:
        if tasks_temp_dir_base.exists() and not any(tasks_temp_dir_base.iterdir()):
            tasks_temp_dir_base.rmdir()
    except OSError as e: logger.warning(f"Could not cleanup base task temp directory {tasks_temp_dir_base}: {e}", exc_info=DEBUG_MODE)


def main():
    global DEBUG_MODE
    parser = argparse.ArgumentParser(
        description="Convert EML/MBOX files to PDF/HTML/Markdown with attachment deduplication and merging options.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )
    # ... (all other arguments as before) ...
    parser.add_argument("--input", required=True, help="Directory containing .eml or .mbox files.")
    parser.add_argument("--outputdir", required=True, help="Directory for generated output files.")
    parser.add_argument("--format", choices=["pdf", "html", "md"], default="pdf", help="Output format.")
    parser.add_argument(
        "--deduplicate-attachments",
        action=argparse.BooleanOptionalAction, # Provides --deduplicate-attachments and --no-deduplicate-attachments
        default=True,
        help="Enable content-based SHA256 deduplication for PDF attachments. Duplicates (within the same output file/part) will be referenced instead of embedded."
    )
    parser.add_argument("--merge", action="store_true", help="Merge all PDFs into parts if --format is pdf.")
    parser.add_argument("--splitsize", type=float, default=0, help="If --merge, split merged PDF parts after this size (MB) of email body PDFs.")
    parser.add_argument("--splitpages", type=int, default=0, help="If --merge, split merged PDF parts after this many pages.")
    parser.add_argument("--report", action="store_true", help="Write a 'eml_conversion_report.txt' summary in output_dir.")
    parser.add_argument("--prepend-date", action="store_true", help="Prepend the email's date (YYYY-MM-DD) to each file's name.")
    parser.add_argument("--mbox", action="store_true", help="Treat input files as mbox format.")
    parser.add_argument("--threads", type=int, default=8, help="Number of worker threads for processing.")
    parser.add_argument("--loglevel", choices=["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"], default="INFO", help="Set the logging output level.")
    args = parser.parse_args()

    # --- Logging Setup ---
    # ... (as before) ...
    log_level_map = {"DEBUG": logging.DEBUG, "INFO": logging.INFO, "WARNING": logging.WARNING, "ERROR": logging.ERROR, "CRITICAL": logging.CRITICAL}
    chosen_log_level = log_level_map.get(args.loglevel.upper(), logging.INFO)
    DEBUG_MODE = (chosen_log_level == logging.DEBUG)
    logging.basicConfig(level=chosen_log_level, format='%(asctime)s - %(levelname)-8s - %(threadName)-12s - %(filename)s:%(lineno)d - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    logging.getLogger('weasyprint').setLevel(logging.WARNING if not DEBUG_MODE else logging.DEBUG)
    logging.getLogger('fontTools').setLevel(logging.WARNING if not DEBUG_MODE else logging.INFO)
    logging.getLogger('PyPDF2').setLevel(logging.WARNING if not DEBUG_MODE else logging.DEBUG)


    if args.merge and args.format != "pdf": parser.error("--merge can only be used with --format pdf")
    if (args.splitsize > 0 or args.splitpages > 0) and not (args.merge and args.format == "pdf"): parser.error("--splitsize and --splitpages can only be used with --merge and --format pdf")
    if args.threads < 1: args.threads = 1; logger.warning("Number of threads set to 1 (minimum).")

    logger.info(f"Starting EML/MBOX conversion process with arguments: {args}")
    convert_eml_files_in_directory(
        eml_dir=args.input, output_dir=args.outputdir, merge_all=args.merge, report=args.report,
        prepend_date=args.prepend_date, out_format=args.format, mbox_mode=args.mbox,
        splitsize=args.splitsize, splitpages=args.splitpages, num_threads=args.threads,
        deduplicate_attachments=args.deduplicate_attachments
    )
    logger.info("EML/MBOX conversion process finished.")

if __name__ == "__main__":
    main()