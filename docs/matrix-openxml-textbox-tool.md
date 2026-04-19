# Matrix OpenXML Textbox Tool

This is a C# helper for Matrix Media DOCX rewrites.

Location:

- [`dotnet/MatrixOpenXmlTextboxTool`](../dotnet/MatrixOpenXmlTextboxTool)

## What It Does

- Opens a `.docx` as an Open XML package
- Scans `word/*.xml` parts
- Rewrites Matrix Media table amount cells using the same non-Oneonta and Oneonta formulas as the old COM path
- Finds Word textbox content nodes (`w:txbxContent`)
- Applies explicit string replacements inside textbox paragraphs when the replacement is unambiguous

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

When `MATRIX_OPENXML_TEXTBOX_TOOL` points to a `.dll`, the wrapper looks for:

- `MATRIX_OPENXML_DOTNET`
- `dotnet` on `PATH`
- `/mnt/c/Program Files/dotnet/dotnet.exe` from WSL
- `/mnt/c/Program Files (x86)/dotnet/dotnet.exe` from WSL

## Pipeline Usage

The default Matrix rewrite backend now tries the native `.NET` OpenXML rewriter first. The Python path falls back to Word COM only when the native rewriter fails, or when `MATRIX_REWRITER=word` is set.

The standalone textbox post-pass still exists for comparison, but the preferred native command is:

```bash
dotnet MatrixOpenXmlTextboxTool.dll rewrite-matrix-amounts --docx /path/to/matrix.docx
```
