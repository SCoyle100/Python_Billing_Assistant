# Matrix OpenXML Textbox Tool

This is an experimental C# helper for Matrix Media textbox rewrites.

Location:

- [`dotnet/MatrixOpenXmlTextboxTool`](../dotnet/MatrixOpenXmlTextboxTool)

## What It Does

- Opens a `.docx` as an Open XML package
- Scans `word/*.xml` parts
- Finds Word textbox content nodes (`w:txbxContent`)
- Applies explicit string replacements inside textbox paragraphs

## What It Does Not Solve

- It is **not** a Word layout engine
- It does **not** know page numbers
- It does **not** know which textbox belongs to which rendered page
- It can only safely apply unambiguous text replacements

That means it is a candidate for:

- straightforward textbox amount rewrites

It is not yet a guaranteed replacement for the current COM textbox flow when:

- the same original amount appears on multiple pages but should map to different rewritten amounts
- page-aware behavior like Oneonta-specific multiplier selection is required for textboxes

## Python Integration

The Python wrapper is:

- [`vendor_invoice_logic/matrix_textbox_openxml.py`](../vendor_invoice_logic/matrix_textbox_openxml.py)

Set:

- `MATRIX_OPENXML_TEXTBOX_TOOL=/absolute/path/to/MatrixOpenXmlTextboxTool.dll`

or:

- `MATRIX_OPENXML_TEXTBOX_TOOL=/absolute/path/to/MatrixOpenXmlTextboxTool`

If the variable is unset, the wrapper does nothing.

## Planned Usage

Once built, the helper is intended to be used only when Matrix textbox replacements are unambiguous. Otherwise the Python path should continue to fall back to COM.

At the moment it is **scaffolded but not enabled in the live rewrite path by default**. The current COM flow keeps the document open inside Word while editing, and an external Open XML post-processor should only be tested against a closed/saved document to avoid overwrite risk.
