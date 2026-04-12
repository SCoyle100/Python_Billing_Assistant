# Linux/macOS Port Plan

Base branch: `remove-dspy-openai-pdf-invoice-seed`  
Planning branch: `linux-macos-port-plan`

## Goal

Make the invoice pipeline run on Linux/macOS, with Linux as the primary target, by removing Microsoft Word COM automation from the production path where possible.

The preferred direction is:

1. Stay in Python if practical.
2. Use a cross-platform renderer/converter only where document layout export is unavoidable.
3. Use C#/.NET only if Python plus LibreOffice is not reliable enough for the remaining fidelity requirements.

## Current State

The latest non-`main` branch is already partway there:

- PDF -> DOCX conversion is already cross-platform in principle because [`pdf_to_docx_.py`](../pdf_to_docx_.py) uses Adobe PDF Services, not Word automation.
- Capitol Media logic is already trending away from fine-grained Word formatting because the table is rebuilt from extracted data.
- Excel is not a current blocker in the main pipeline. I did not find active Excel COM automation in production code.
- [`pdf_to_docx_.py`](../pdf_to_docx_.py) still imports `win32com` and PyQt even though its active method only performs Adobe PDF Services conversion. That is an easy early cleanup because the unused import can still break Linux startup.

The main blockers are all on the Word/DOCX/PDF side.

## What COM Is Doing Right Now

Production path usage:

- [`main_logic.py`](../main_logic.py) calls the COM-backed Matrix path:
  - `read_page_markets()` at lines 640-641
  - `analyze_word_document()` at lines 642-643
  - `build_dataframe_from_word_document()` at line 646
  - `create_images_from_docx()` at lines 717-722 and 778
- [`vendor_invoice_logic/matrix_media_market_map.py`](../vendor_invoice_logic/matrix_media_market_map.py) opens Word and uses page information from the Word object model to map page -> `(market, service_period)`.
- [`vendor_invoice_logic/matrix_media_logic.py`](../vendor_invoice_logic/matrix_media_logic.py) opens Word, edits amount cells in-place, and also updates floating text boxes/shapes on the page.
- [`vendor_invoice_logic/matrix_media_dataframe.py`](../vendor_invoice_logic/matrix_media_dataframe.py) opens Word to read Matrix tables and extract rows.
- [`vendor_invoice_logic/capitol_media_dataframe_1.py`](../vendor_invoice_logic/capitol_media_dataframe_1.py) still opens Word first to identify the invoice table, even though the deterministic extraction step already uses `python-docx`.
- [`image_generation/create_pdf_image.py`](../image_generation/create_pdf_image.py) uses Word COM to save DOCX -> PDF before converting PDF pages to PNG.

Non-production / utility usage:

- [`pdf_to_docx.py`](../pdf_to_docx.py) contains legacy/manual Word automation helpers.
- [`gen_track.py`](../gen_track.py) inspects the pywin32 COM cache path.

## Other Platform-Sensitive Things

These are not COM, but they still matter:

- [`invoice_processor.py`](../invoice_processor.py) hardcodes `D:\\Tesseract\\tesseract.exe`. That is Windows-specific and will fail on Linux/macOS as written.
- PyQt5 and tkinter file dialogs are used in several entry points. Those are cross-platform GUI libraries, not Windows-native, but they are not suitable for a headless Linux deployment.
- There are Windows-style paths in utility/debug code, but I did not find active Excel interop or other Win32 APIs in the main invoice pipeline.

## Recommendation

Python-first is realistic here.

I would **not** start with a C# rewrite.

Reasoning:

- `python-docx` is already used widely in this repo and is a good fit for reading/writing tables and rebuilding simple invoice layouts.
- PyDocX is **not** the right replacement for COM in this project. Its documented use is DOCX -> HTML conversion, not editing DOCX files or rendering them to PDF.
- The hardest gap is not “can we parse Word?” It is “how do we render DOCX to PDF cross-platform with acceptable fidelity?”
- A C# move to Open XML SDK would help with document structure manipulation, but it would not solve pagination/PDF conversion/rendering by itself.

## Library Assessment

### Good Python-first choices

- `python-docx`
  - Use for table extraction, deterministic row parsing, and rebuilding simplified DOCX output.
- `pandas`
  - Already sufficient for the existing Excel read path.
- `openpyxl`
  - Add only if you need direct workbook-level reads/writes beyond `pandas.read_excel()`.
- `PyMuPDF` / `pdfplumber`
  - Use to read text per PDF page and build page -> market mappings without asking Word for page numbers.
- LibreOffice CLI (`soffice --headless --convert-to ...`)
  - Best practical cross-platform DOCX/XLSX -> PDF backend for Linux/macOS if you want to stay out of commercial libraries.

### Not a good primary replacement here

- PyDocX
  - Useful for DOCX -> HTML, but not for preserving/rebuilding the Word pipeline you currently have.
- `docx2pdf`
  - Not usable for Linux. Its own README says it converts via Microsoft Word on Windows or macOS only.

### C#/.NET fallback, only if needed

- Open XML SDK
  - Good for manipulating DOCX/XLSX structure.
  - Not enough on its own for repagination/PDF rendering.
- Aspose.Words / Syncfusion DocIO
  - These are the class of libraries to consider if you need cross-platform rendering with higher fidelity and are willing to accept commercial licensing.

## Proposed Rewrite Plan

### Phase 1: isolate platform-specific behavior

Create explicit backends/interfaces around:

- PDF -> DOCX conversion
- DOCX parsing
- page -> market mapping
- amount rewrite/rebuild
- DOCX -> PDF rendering
- PDF -> PNG generation

Deliverable:

- One thin orchestration path in `main_logic.py` that depends on backend functions rather than directly on COM-backed modules.

### Phase 2: remove low-risk COM first

Replace Capitol Media table discovery in `vendor_invoice_logic/capitol_media_dataframe_1.py` with pure `python-docx` table scanning.

Expected result:

- Capitol Media should become COM-free first.
- This is the lowest-risk place to prove the Linux direction.

### Phase 3: replace Matrix data extraction

Rewrite `vendor_invoice_logic/matrix_media_dataframe.py` to use `python-docx` instead of Word COM for table traversal.

Expected result:

- Market/amount/service period/description extraction becomes cross-platform.
- Fort Payne normalization logic stays unchanged.

### Phase 4: replace Matrix page mapping

Do **not** try to reproduce Word pagination in Python.

Instead, derive page mapping from the source PDF, not the converted DOCX:

- Read each PDF page with `PyMuPDF` or `pdfplumber`
- Extract the first market/service-period pair per page
- Normalize market names the same way the current code does

Expected result:

- Remove dependence on `table.Range.Information(wdActiveEndPageNumber)`.
- Align image naming with the actual source PDF pages, which is what you finally rasterize anyway.

### Phase 5: replace Matrix amount rewrite

Two options:

1. Preferred: rebuild the relevant table content rather than editing Word’s original layout in place.
2. Fallback: edit table cells directly with `python-docx` and accept that floating text boxes/shapes may need a separate strategy.

My recommendation is to favor rebuild logic, similar to what you already did for Capitol Media, if the business requirement is invoice correctness rather than pixel-perfect source formatting.

Risk:

- The current COM code also updates floating shapes/text boxes. `python-docx` does not provide equivalent support for those objects in its normal API.

### Phase 6: replace DOCX -> PDF rendering

Introduce a renderer abstraction with at least:

- `LibreOfficeRenderer` for Linux/macOS via `soffice --headless --convert-to`
- `WordComRenderer` as an optional Windows fallback during migration

Expected result:

- Same high-level pipeline on all platforms
- Windows-only behavior isolated behind one backend instead of spread across the codebase

### Phase 7: headless Linux run mode

Separate GUI/manual tools from the production batch path:

- move file selection dialogs behind optional CLI flags or a separate manual tools package
- make the main pipeline runnable without PyQt or tkinter

Expected result:

- CI-friendly and server-friendly execution on Linux

### Phase 8: regression fixtures

Add sample fixtures for at least:

- Matrix Media with Fort Payne
- Matrix Media with Oneonta
- Capitol Media rebuilt table flow
- a fee-invoice-only run

Verify:

- extracted invoices
- service periods
- invoice-image naming
- generated PDF/image counts

## Suggested Order Of Work

1. Capitol Media COM removal
2. Matrix extraction COM removal
3. Matrix page mapping from PDF pages
4. DOCX -> PDF renderer abstraction with LibreOffice
5. Matrix table rebuild strategy
6. Headless CLI cleanup and tests

## My Read On The Best Path

The repo does **not** need a broad C# rewrite to get to Linux.

The most pragmatic path is:

- keep Python for orchestration, parsing, DB work, and invoice generation
- replace COM usage with `python-docx` where the task is structural document editing/parsing
- use PDF-page text extraction for page-aware mapping
- use LibreOffice as the cross-platform rendering backend for DOCX/XLSX -> PDF

Only move to C# if LibreOffice rendering fidelity becomes unacceptable or if you need a single supported commercial engine for both editing and rendering.

## Source Notes

External references used for the recommendations above:

- `python-docx` tables and shapes docs:
  - https://python-docx.readthedocs.io/en/latest/user/tables.html
  - https://python-docx.readthedocs.io/en/latest/user/shapes.html
- PyDocX usage docs:
  - https://pydocx.readthedocs.io/en/latest/usage.html
- `docx2pdf` README:
  - https://github.com/AlJohri/docx2pdf
- LibreOffice CLI docs:
  - https://help.libreoffice.org/latest/ug/text/shared/guide/start_parameters.html
  - https://help.libreoffice.org/latest/sid/text/shared/guide/pdf_params.html
- Open XML SDK and rendering limitations context:
  - https://learn.microsoft.com/office/open-xml/about-the-open-xml-sdk
  - https://learn.microsoft.com/en-us/previous-versions/office/developer/sharepoint-2010/ff742315(v=office.14)
- Commercial .NET fallback examples:
  - https://docs.aspose.com/words/net/system-requirements/
  - https://help.syncfusion.com/document-processing/system-requirements
