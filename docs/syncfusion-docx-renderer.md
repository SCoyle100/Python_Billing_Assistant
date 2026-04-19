# Syncfusion DOCX Renderer

This branch can use Syncfusion DocIO to convert DOCX files to PDF without Microsoft Word.

## NuGet Packages

The .NET document helper references:

- `Syncfusion.DocIORenderer.Net.Core`
- `SkiaSharp.NativeAssets.Linux`
- `HarfBuzzSharp.NativeAssets.Linux`

These are installed in:

- [`dotnet/MatrixOpenXmlTextboxTool`](../dotnet/MatrixOpenXmlTextboxTool)

## License

Set the license key in the environment before running the pipeline:

```bash
export SYNCFUSION_LICENSE_KEY="..."
```

Do not commit the actual key. The code reads it from `SYNCFUSION_LICENSE_KEY` and registers it inside the .NET helper before running Syncfusion conversion.

## Tool Command

Build:

```bash
dotnet build dotnet/MatrixOpenXmlTextboxTool/MatrixOpenXmlTextboxTool.csproj -c Release
```

Convert:

```bash
dotnet dotnet/MatrixOpenXmlTextboxTool/bin/Release/net8.0/MatrixOpenXmlTextboxTool.dll \
  convert-docx-to-pdf \
  --docx /path/to/input.docx \
  --output /path/to/output.pdf
```

Rewrite Matrix amounts without Word COM:

```bash
dotnet dotnet/MatrixOpenXmlTextboxTool/bin/Release/net8.0/MatrixOpenXmlTextboxTool.dll \
  rewrite-matrix-amounts \
  --docx /path/to/input.docx
```

From WSL, the Python wrapper can also discover the Windows .NET SDK at:

- `/mnt/c/Program Files/dotnet/dotnet.exe`

## Python Integration

The Python wrapper is:

- [`document_backends/dotnet_document_tool.py`](../document_backends/dotnet_document_tool.py)

The default Matrix rewrite backend tries the .NET OpenXML rewriter first and falls back to Word COM only if the native rewriter fails. Set `MATRIX_REWRITER=word` to force the old COM path for comparison.

When `SYNCFUSION_LICENSE_KEY` is set, the default rendering backend uses Syncfusion for DOCX to PDF, then uses the existing PyMuPDF rasterization path for PDF to PNG.
