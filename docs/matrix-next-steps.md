# Matrix Next Steps

This branch now includes:

- the Python backend seam for portability work
- a PDF-first Matrix page-mapper with COM fallback
- a scaffolded C# OpenXML textbox helper

The C# helper is **not** wired into the live rewrite path yet.

## 1. Build The Helper

Build:

- [`dotnet/MatrixOpenXmlTextboxTool`](../dotnet/MatrixOpenXmlTextboxTool)

Expected command once `.NET` is installed:

```bash
dotnet build dotnet/MatrixOpenXmlTextboxTool/MatrixOpenXmlTextboxTool.csproj -c Release
```

After that, point the Python wrapper at the built executable or `.dll`:

- `MATRIX_OPENXML_TEXTBOX_TOOL=/absolute/path/to/MatrixOpenXmlTextboxTool.dll`

or:

- `MATRIX_OPENXML_TEXTBOX_TOOL=/absolute/path/to/MatrixOpenXmlTextboxTool`

Relevant wrapper:

- [`vendor_invoice_logic/matrix_textbox_openxml.py`](../vendor_invoice_logic/matrix_textbox_openxml.py)

## 2. Run A Closed-Document Post-Pass Test On A Real Matrix Sample

What this means:

- take a real Matrix Media `.pdf`
- let the current pipeline convert/process it up through the normal Word COM rewrite
- make sure Word has fully saved and closed the resulting `.docx`
- only **after that**, run the C# OpenXML helper against that saved `.docx`

Why “closed-document” matters:

- if Word still has the `.docx` open, an external OpenXML editor can make changes that Word later overwrites when it saves
- the OpenXML helper should operate on a file that is no longer open in Word

Why “real Matrix sample” matters:

- the textbox problem is page/layout/vendor-template sensitive
- a synthetic test is not enough to prove the real document structure matches what the helper expects
- we need to see whether the actual Matrix textboxes store amounts in a way that the helper can replace safely

What to verify in that test:

1. The normal table amounts are still correct after the existing COM rewrite.
2. The helper changes the expected textbox amount text.
3. It does not damage unrelated text runs or document structure.
4. The resulting `.docx` still opens correctly in Word.
5. The generated PDF/images still look correct after any later pipeline steps.

## 3. Wire It Into The Pipeline After Word Save/Close

What this means:

- keep the current COM-based Matrix rewrite step for now
- let it finish editing tables and any COM-managed content
- save the document
- close the document in Word
- quit Word
- then invoke the OpenXML textbox helper as a separate post-processing step on the saved `.docx`

That is different from the unsafe version:

- unsafe: call the OpenXML helper while Word still has the file open
- safer: call the OpenXML helper only after the COM phase has finished and released the file

Likely integration point:

- [`vendor_invoice_logic/matrix_media_logic.py`](../vendor_invoice_logic/matrix_media_logic.py)

The practical flow would be:

1. COM opens Matrix `.docx`
2. COM rewrites table amounts
3. COM saves and closes Word document
4. Python invokes the OpenXML helper on the saved `.docx`
5. Later pipeline steps continue, such as PDF rendering or image generation

## Recommendation

Do not replace the COM textbox path yet.

The next safe move is:

1. build the helper
2. run a real-sample post-pass test
3. only then decide whether to:
   - keep COM as primary
   - use OpenXML for a subset of safe textbox rewrites
   - wait for Syncfusion and test a stronger rendering/editing engine
