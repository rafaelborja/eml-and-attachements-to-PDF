EML/MBOX Converter (to PDF, HTML, MD)
=====================================

A Python command-line tool to convert `.eml` and `.mbox`/`.mbx` (email) files to PDF, HTML, or Markdown. It's designed to run on Windows, Linux, macOS, and other platforms where Python and its dependencies can be installed.

This tool has been tested with Google Takeout MBOX files.

Key Features
------------

-   Versatile Input: Processes individual `.eml` files or messages within `.mbox` (and `.mbx`) archives.

-   Multiple Output Formats: Converts emails to `.pdf`, `.html`, or `.md` files.

-   Content Extraction: Extracts HTML or text content from each email.

-   Inline Image Embedding: For PDF and HTML outputs, embeds inline images (`cid:` references) by base64-encoding them.

-   Comprehensive Attachment Handling (PDF Output):

   -   Appends PDF attachments to the main email's PDF.

   -   Each attachment (original or reference) is preceded by an information page showing its filename and SHA256 content hash.

   -   Content-Based Deduplication: Identifies duplicate PDF attachments based on their SHA256 hash.

      -   The first instance of an attachment is embedded fully.

      -   Subsequent identical attachments (within the same output PDF part) are referenced with a notice, saving space. This is enabled by default (`--deduplicate-attachments`).

   -   Handles corrupted/unreadable PDF attachments by inserting a notice page instead of crashing.

-   Header Display: Shows email headers (Date, From, To, Cc, Subject) at the top of the output.

-   Multithreading: Processes multiple emails in parallel for significantly faster conversions, especially for large batches (`--threads` option).

-   Advanced PDF Merging:

   -   Optionally merges all generated PDFs into final output files (`--merge`).

   -   Supports splitting these merged PDF files by size (`--splitsize`) or page count (`--splitpages`).

-   Filename Customization: Optionally prepends the email's date (YYYY-MM-DD) to output filenames (`--prepend-date`).

-   Reporting & Logging:

   -   Optionally generates a text-based summary report (`eml_conversion_report.txt`) (`--report`).

   -   Configurable logging levels for detailed diagnostics (`--loglevel`).

1\. Installation
----------------

1.  Install Python 3.9+.

2.  WeasyPrint Dependencies: For PDF output, WeasyPrint is used.

   -   On Linux/macOS, typical dependencies are Pango, Cairo, and GDK-PixBuf. Install via your system's package manager (e.g., `apt-get install libpango-1.0-0 libcairo2 libgdk-pixbuf2.0-0` on Debian/Ubuntu).

   -   On Windows, install the GTK3 runtime environment as per the [WeasyPrint documentation](https://doc.courtbouillon.org/weasyprint/stable/first_steps.html#windows "null"). (Using MSYS2 to install dependencies is also an option).

3.  Create and activate a virtual environment (recommended):

    ```
    python -m venv venv
    # On Linux/macOS:
    source venv/bin/activate
    # On Windows (PowerShell):
    # .\venv\Scripts\Activate.ps1
    # On Windows (CMD):
    # .\venv\Scripts\activate.bat

    ```

4.  Install required Python libraries:

    ```
    pip install weasyprint PyPDF2 html2text

    ```

5.  Verify WeasyPrint installation (optional, but good for PDF troubleshooting):

    ```
    python -m weasyprint --info

    ```

    This should output version information for WeasyPrint and its dependencies.

2\. Usage
---------

The script is run from the command line. Let's assume your script is named `eml_converter.py`.

```
python eml_converter.py --input <INPUT_DIRECTORY> --outputdir <OUTPUT_DIRECTORY> [OPTIONS...]

```

### Command-Line Options

-   `--input <DIR>`: (Required) Directory containing `.eml` or `.mbox`/`.mbx` files.

-   `--outputdir <DIR>`: (Required) Directory for generated output files.

-   `--format {pdf,html,md}`: Output format. (Default: `pdf`)

-   `--deduplicate-attachments` / `--no-deduplicate-attachments`: Enable/disable content-based SHA256 deduplication for PDF attachments within the same output file/part. (Default: enabled)

-   `--merge`: Merge all generated PDFs into final output parts (applies only if `format` is `pdf`).

-   `--splitsize <MB>`: If `--merge`, split merged PDF parts after this size (MB) of combined email body PDFs. (Default: `0` - no size split)

-   `--splitpages <NUM>`: If `--merge`, split merged PDF parts after this many pages (sum of body pages and estimated attachment pages). (Default: `0` - no page split)

-   `--report`: Write a summary report (`eml_conversion_report.txt`) in the output directory.

-   `--prepend-date`: Prepend the email's date (YYYY-MM-DD) to output filenames.

-   `--mbox`: Treat input files as MBOX format (searches for `.mbox` or `.mbx` extensions in the input directory). If not set, treats files as individual `.eml` files.

-   `--threads <NUM>`: Number of worker threads for processing. (Default: `8`)

-   `--loglevel {DEBUG,INFO,WARNING,ERROR,CRITICAL}`: Set the logging output level. (Default: `INFO`)

### Examples

```
# Basic: Convert EMLs in 'emails/' to individual PDFs in 'output_pdfs/'
python eml_converter.py --input ./emails --outputdir ./output_pdfs

# MBOX: Process MBOX files in 'archives/' to HTML, 4 threads
python eml_converter.py --input ./archives --outputdir ./output_html --mbox --format html --threads 4

# Advanced PDF: Merge all, split by 50MB, prepend date, deduplicate (default), generate report
python eml_converter.py --input ./eml_files --outputdir ./pdf_merged --merge --splitsize 50 --prepend-date --report

# Disable attachment deduplication for PDF output
python eml_converter.py --input ./eml_files --outputdir ./pdf_output --no-deduplicate-attachments

# Debugging: Process with detailed logs
python eml_converter.py --input ./emails --outputdir ./debug_out --loglevel DEBUG

```

3\. Output
----------

-   Individual Files: For each processed email, an output file (e.g., `.pdf`, `.html`, `.md`) is created in the specified `<OUTPUT_DIRECTORY>`.

-   Merged PDFs: If `--merge` and `--format pdf` are used:

   -   Files will be named `merged_output_part_XXX.pdf`.

   -   If no splitting options are used, a single `merged_output_part_001.pdf` will typically be produced.

-   Report: If `--report` is used, an `eml_conversion_report.txt` file is generated in the output directory, summarizing the conversion process, including counts of processed items, failures, attachments, and reused attachments (if deduplication was active for PDF).

4\. Attachment Deduplication (PDF Output)
-----------------------------------------

When processing emails to PDF and attachment deduplication is active (default):

1.  The content of each PDF attachment is hashed using SHA256.

2.  An information page is generated for *every* PDF attachment, displaying its original filename and its SHA256 hash (Content ID).

3.  If an attachment's content (based on its hash) is encountered for the first time within the current output PDF file (or merged part), its content is embedded after its information page.

4.  If the same content (identical hash) is encountered again within the same output PDF file (or merged part):

   -   The information page will include a notice: `DUPLICATE` CONTENT: This attachment ('current_file.pdf') has the exact same content as attachment 'first_occurrence.pdf' `(ID: <sha256_hash>) previously in this document part. The content is not repeated here.`

   -   The actual attachment content is *not* embedded again for this instance.

5.  This deduplication scope is reset for each new file created by the splitting logic (`--splitsize`, `--splitpages`).

This feature helps reduce the size of the output PDFs when emails contain many identical attachments.

5\. Handling Broken Attachments (PDF Output)
--------------------------------------------

If a PDF attachment is corrupted or cannot be processed by `PyPDF2`:

-   The script catches exceptions during attachment processing.

-   A warning is logged.

-   The standard information page for that attachment will be generated.

-   If the content embedding fails (e.g., for a corrupted PDF), a placeholder page or message indicating the failure may be included instead of the attachment's content.

-   The filename of the problematic attachment is recorded and listed in the summary report.

6\. Project Structure
---------------------

A typical layout after running might include:

```
eml_converter.py        # Your script
requirements.txt        # If you create one: weasyprint, PyPDF2, html2text
README.md               # This file
venv/                   # Virtual environment (recommended)
tmp_eml_attachments/    # Auto-created: Temporary WeasyPrint artifacts (e.g., header pages)
tmp_task_processing/    # Auto-created: Temporary intermediate files (e.g., body PDFs for merging)

```

The `tmp_...` directories are used for transient files during processing and may be cleaned up.

7\. Changelog
-------------

-   Recent Major Updates (leading up to June 2025):

   -   Added MBOX (.mbox, .mbx) file processing via `--mbox` flag.

   -   Introduced multithreading (`--threads`) for significantly faster processing of multiple emails.

   -   Added support for HTML (`--format html`) and Markdown (`--format md`) output formats.

   -   Implemented SHA256 content-based deduplication for PDF attachments (`--deduplicate-attachments`, default on). Duplicates within the same output PDF part are referenced instead of re-embedded.

   -   Enhanced PDF merging with splitting by size (`--splitsize`) or page count (`--splitpages`).

   -   Added option to prepend email date to output filenames (`--prepend-date`).

   -   Integrated configurable logging levels (`--loglevel`) for better diagnostics.

   -   Refined temporary file handling and overall workflow for robustness.

   -   Summary report now includes count of reused (deduplicated) attachments.

-   Initial Version:

   -   Basic EML to PDF conversion.

   -   Inline image embedding.

   -   PDF attachment appending.

   -   Header display.

   -   Optional merging into a single PDF.

   -   Basic summary report.

8\. License
-----------

(C) 2025 Rafael Borja

Distributed under the Apache 2.0 License. (Assuming you have a LICENSE file or intend to use Apache 2.0)

### Final Notes

-   MIME-Encoded Headers: The script decodes MIME-encoded headers (e.g., for subjects, sender names with special characters) using Python's `email.header.decode_header` for proper display in the output.

-   WeasyPrint Dependencies: Ensure WeasyPrint's underlying dependencies (Pango, Cairo, etc.) are correctly installed on your system, especially on Windows, for PDF generation to work reliably.

-   Temporary Directories: The script creates `tmp_eml_attachments` and `tmp_task_processing` for intermediate files. These are generally cleaned up, but if the script is interrupted, some temporary files might remain.