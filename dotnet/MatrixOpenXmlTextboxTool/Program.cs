using System.IO.Compression;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocIORenderer;
using Syncfusion.Licensing;
using Syncfusion.Pdf;

if (args.Length == 0)
{
    PrintUsage();
    return 1;
}

return args[0] switch
{
    "convert-docx-to-pdf" => ConvertDocxToPdf(args.Skip(1).ToArray()),
    "rewrite-matrix-amounts" => RewriteMatrixAmounts(args.Skip(1).ToArray()),
    "rewrite-textboxes" => RewriteTextboxes(args.Skip(1).ToArray()),
    _ => Fail($"Unknown command '{args[0]}'.")
};

static int RewriteMatrixAmounts(string[] args)
{
    string? docxPath = null;

    for (var index = 0; index < args.Length; index++)
    {
        switch (args[index])
        {
            case "--docx":
                docxPath = NextValue(args, ref index, "--docx");
                break;
            default:
                return Fail($"Unknown argument '{args[index]}'.");
        }
    }

    if (string.IsNullOrWhiteSpace(docxPath))
    {
        return Fail("--docx is required.");
    }

    if (!File.Exists(docxPath))
    {
        return Fail($"DOCX file not found: {docxPath}");
    }

    var textboxReplacements = new Dictionary<string, string>(StringComparer.Ordinal);
    var ambiguousTextboxAmounts = new HashSet<string>(StringComparer.Ordinal);
    var modifiedTableCells = 0;
    var tableReplacementCount = 0;
    var modifiedTextboxParagraphs = 0;
    var textboxReplacementCount = 0;

    try
    {
        using var archive = ZipFile.Open(docxPath, ZipArchiveMode.Update);
        var loadedParts = archive.Entries
            .Where(entry =>
                entry.FullName.StartsWith("word/", StringComparison.OrdinalIgnoreCase) &&
                entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .Select(entry => TryLoadXml(entry, out var document) && document?.Root is not null
                ? new OpenXmlPartDocument(entry, document)
                : null)
            .Where(item => item is not null)
            .Cast<OpenXmlPartDocument>()
            .ToList();

        foreach (var part in loadedParts)
        {
            if (ProcessMatrixTablesInPart(
                part.Document,
                textboxReplacements,
                ambiguousTextboxAmounts,
                out var partModifiedCells,
                out var partTableReplacements))
            {
                part.Changed = true;
                modifiedTableCells += partModifiedCells;
                tableReplacementCount += partTableReplacements;
            }
        }

        var safeTextboxReplacements = textboxReplacements
            .Where(item => !ambiguousTextboxAmounts.Contains(item.Key))
            .Select(item => new ReplacementRule { Original = item.Key, Replacement = item.Value })
            .ToList();

        if (safeTextboxReplacements.Count > 0)
        {
            foreach (var part in loadedParts)
            {
                if (ProcessOpenXmlPart(
                    part.Document,
                    safeTextboxReplacements,
                    out var partModifiedTextboxes,
                    out var partTextboxReplacements))
                {
                    part.Changed = true;
                    modifiedTextboxParagraphs += partModifiedTextboxes;
                    textboxReplacementCount += partTextboxReplacements;
                }
            }
        }

        foreach (var part in loadedParts.Where(part => part.Changed))
        {
            SaveXml(part.Entry, part.Document);
        }
    }
    catch (Exception exc)
    {
        return Fail($"Matrix Open XML rewrite failed: {exc.Message}");
    }

    if (ambiguousTextboxAmounts.Count > 0)
    {
        Console.Error.WriteLine(
            "Skipped ambiguous textbox amount(s): " + string.Join(", ", ambiguousTextboxAmounts.Order()));
    }

    Console.WriteLine(
        $"Updated {tableReplacementCount} table amount replacement(s) across {modifiedTableCells} cell(s).");
    Console.WriteLine(
        $"Updated {textboxReplacementCount} textbox replacement(s) across {modifiedTextboxParagraphs} paragraph(s).");

    return 0;
}

static int ConvertDocxToPdf(string[] args)
{
    string? docxPath = null;
    string? outputPath = null;

    for (var index = 0; index < args.Length; index++)
    {
        switch (args[index])
        {
            case "--docx":
                docxPath = NextValue(args, ref index, "--docx");
                break;
            case "--output":
                outputPath = NextValue(args, ref index, "--output");
                break;
            default:
                return Fail($"Unknown argument '{args[index]}'.");
        }
    }

    if (string.IsNullOrWhiteSpace(docxPath) || string.IsNullOrWhiteSpace(outputPath))
    {
        return Fail("Both --docx and --output are required.");
    }

    if (!File.Exists(docxPath))
    {
        return Fail($"DOCX file not found: {docxPath}");
    }

    if (!RegisterSyncfusionLicense())
    {
        return 1;
    }

    try
    {
        var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        using var inputStream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var wordDocument = new WordDocument(inputStream, FormatType.Docx);
        using var renderer = new DocIORenderer();
        using PdfDocument pdfDocument = renderer.ConvertToPDF(wordDocument);
        using var outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
        pdfDocument.Save(outputStream);

        Console.WriteLine($"Saved PDF: {outputPath}");
        return 0;
    }
    catch (Exception exc)
    {
        return Fail($"DOCX to PDF conversion failed: {exc.Message}");
    }
}

static bool RegisterSyncfusionLicense()
{
    var licenseKey = Environment.GetEnvironmentVariable("SYNCFUSION_LICENSE_KEY");
    if (string.IsNullOrWhiteSpace(licenseKey))
    {
        Console.Error.WriteLine("SYNCFUSION_LICENSE_KEY is required for Syncfusion document conversion.");
        return false;
    }

    SyncfusionLicenseProvider.RegisterLicense(licenseKey);
    return true;
}

static int RewriteTextboxes(string[] args)
{
    string? docxPath = null;
    string? replacementsPath = null;

    for (var index = 0; index < args.Length; index++)
    {
        switch (args[index])
        {
            case "--docx":
                docxPath = NextValue(args, ref index, "--docx");
                break;
            case "--replacements":
                replacementsPath = NextValue(args, ref index, "--replacements");
                break;
            default:
                return Fail($"Unknown argument '{args[index]}'.");
        }
    }

    if (string.IsNullOrWhiteSpace(docxPath) || string.IsNullOrWhiteSpace(replacementsPath))
    {
        return Fail("Both --docx and --replacements are required.");
    }

    if (!File.Exists(docxPath))
    {
        return Fail($"DOCX file not found: {docxPath}");
    }

    if (!File.Exists(replacementsPath))
    {
        return Fail($"Replacement file not found: {replacementsPath}");
    }

    ReplacementPayload? payload;
    try
    {
        payload = JsonSerializer.Deserialize<ReplacementPayload>(
            File.ReadAllText(replacementsPath),
            new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
    }
    catch (Exception exc)
    {
        return Fail($"Failed to parse replacements JSON: {exc.Message}");
    }

    var replacements = payload?.Replacements?
        .Where(item => !string.IsNullOrWhiteSpace(item.Original) && item.Original != item.Replacement)
        .ToList() ?? new List<ReplacementRule>();

    if (replacements.Count == 0)
    {
        Console.WriteLine("No replacements supplied.");
        return 0;
    }

    var modifiedParts = 0;
    var modifiedTextboxes = 0;
    var replacementCount = 0;

    try
    {
        using var archive = ZipFile.Open(docxPath, ZipArchiveMode.Update);
        var candidateEntries = archive.Entries
            .Where(entry =>
                entry.FullName.StartsWith("word/", StringComparison.OrdinalIgnoreCase) &&
                entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var entry in candidateEntries)
        {
            if (!TryLoadXml(entry, out var document))
            {
                continue;
            }

            if (document?.Root is null)
            {
                continue;
            }

            if (!ProcessOpenXmlPart(document, replacements, out var partTextboxes, out var partReplacements))
            {
                continue;
            }

            SaveXml(entry, document);
            modifiedParts++;
            modifiedTextboxes += partTextboxes;
            replacementCount += partReplacements;
        }
    }
    catch (Exception exc)
    {
        return Fail($"Open XML rewrite failed: {exc.Message}");
    }

    Console.WriteLine(
        $"Updated {replacementCount} replacement(s) across {modifiedTextboxes} textbox paragraph(s) in {modifiedParts} part(s).");
    return replacementCount > 0 ? 0 : 2;
}

static bool ProcessOpenXmlPart(
    XDocument document,
    IReadOnlyList<ReplacementRule> replacements,
    out int modifiedTextboxParagraphs,
    out int replacementCount)
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    modifiedTextboxParagraphs = 0;
    replacementCount = 0;
    var changed = false;

    var textboxes = document.Descendants(w + "txbxContent").ToList();
    foreach (var textbox in textboxes)
    {
        foreach (var paragraph in textbox.Descendants(w + "p"))
        {
            var textNodes = paragraph.Descendants(w + "t").ToList();
            if (textNodes.Count == 0)
            {
                continue;
            }

            var originalText = string.Concat(textNodes.Select(node => node.Value));
            var updatedText = ApplyReplacements(originalText, replacements, out var paragraphReplacementCount);

            if (paragraphReplacementCount == 0 || updatedText == originalText)
            {
                continue;
            }

            textNodes[0].Value = updatedText;
            foreach (var extraNode in textNodes.Skip(1))
            {
                extraNode.Value = string.Empty;
            }

            modifiedTextboxParagraphs++;
            replacementCount += paragraphReplacementCount;
            changed = true;
        }
    }

    return changed;
}

static bool ProcessMatrixTablesInPart(
    XDocument document,
    IDictionary<string, string> textboxReplacements,
    ISet<string> ambiguousTextboxAmounts,
    out int modifiedCells,
    out int replacementCount)
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    modifiedCells = 0;
    replacementCount = 0;
    var changed = false;

    foreach (var table in document.Descendants(w + "tbl").ToList())
    {
        var rows = table.Elements(w + "tr").ToList();
        if (rows.Count < 2)
        {
            continue;
        }

        var headerCells = rows[0].Elements(w + "tc").ToList();
        var marketColumnIndex = FindColumnIndex(headerCells, "Market");
        var amountColumnIndex = FindColumnIndex(headerCells, "Amount");
        if (marketColumnIndex is null || amountColumnIndex is null)
        {
            continue;
        }

        foreach (var row in rows.Skip(1))
        {
            var cells = row.Elements(w + "tc").ToList();
            if (marketColumnIndex.Value >= cells.Count || amountColumnIndex.Value >= cells.Count)
            {
                continue;
            }

            var marketText = GetOpenXmlText(cells[marketColumnIndex.Value]);
            var isOneonta = marketText.Contains("Oneonta", StringComparison.OrdinalIgnoreCase);
            var amountCell = cells[amountColumnIndex.Value];
            var amountText = GetOpenXmlText(amountCell);
            var matches = DollarAmountRegex().Matches(amountText).Cast<Match>().ToList();
            if (matches.Count == 0)
            {
                continue;
            }

            var cellReplacementCount = 0;
            foreach (var match in matches)
            {
                var originalAmount = match.Value;
                var updatedAmount = CalculateUpdatedAmount(originalAmount, isOneonta);
                RecordReplacement(textboxReplacements, ambiguousTextboxAmounts, originalAmount, updatedAmount);
                cellReplacementCount += ReplaceTextInOpenXmlElement(amountCell, originalAmount, updatedAmount);
            }

            if (cellReplacementCount > 0)
            {
                modifiedCells++;
                replacementCount += cellReplacementCount;
                changed = true;
            }
        }
    }

    return changed;
}

static int? FindColumnIndex(IReadOnlyList<XElement> headerCells, string columnName)
{
    for (var index = 0; index < headerCells.Count; index++)
    {
        if (GetOpenXmlText(headerCells[index]).Contains(columnName, StringComparison.OrdinalIgnoreCase))
        {
            return index;
        }
    }

    return null;
}

static string GetOpenXmlText(XElement element)
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    return string.Concat(element.Descendants(w + "t").Select(node => node.Value));
}

static int ReplaceTextInOpenXmlElement(XElement element, string original, string replacement)
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var textNodes = element.Descendants(w + "t").ToList();
    if (textNodes.Count == 0 || string.IsNullOrEmpty(original) || original == replacement)
    {
        return 0;
    }

    var replacementCount = 0;
    foreach (var textNode in textNodes)
    {
        var occurrences = CountOccurrences(textNode.Value, original);
        if (occurrences == 0)
        {
            continue;
        }

        textNode.Value = textNode.Value.Replace(original, replacement, StringComparison.Ordinal);
        replacementCount += occurrences;
    }

    if (replacementCount > 0)
    {
        return replacementCount;
    }

    var originalText = string.Concat(textNodes.Select(node => node.Value));
    var updatedText = originalText.Replace(original, replacement, StringComparison.Ordinal);
    if (updatedText == originalText)
    {
        return 0;
    }

    textNodes[0].Value = updatedText;
    foreach (var extraNode in textNodes.Skip(1))
    {
        extraNode.Value = string.Empty;
    }

    return CountOccurrences(originalText, original);
}

static void RecordReplacement(
    IDictionary<string, string> replacements,
    ISet<string> ambiguousAmounts,
    string originalAmount,
    string updatedAmount)
{
    if (!replacements.TryGetValue(originalAmount, out var existingValue))
    {
        replacements[originalAmount] = updatedAmount;
        return;
    }

    if (existingValue != updatedAmount)
    {
        ambiguousAmounts.Add(originalAmount);
    }
}

static string CalculateUpdatedAmount(string originalAmount, bool isOneonta)
{
    var parsedValue = ParseDollarAmount(originalAmount);
    var updatedValue = isOneonta ? parsedValue * 1.3177 : parsedValue / 0.85;
    var truncatedValue = Math.Abs(updatedValue - Math.Truncate(updatedValue)) > double.Epsilon
        ? Math.Truncate(updatedValue)
        : updatedValue;

    return FormatDollarAmount(truncatedValue);
}

static double ParseDollarAmount(string dollarAmount)
{
    var cleaned = Regex.Replace(dollarAmount, @"[^\d.]", string.Empty);
    return double.TryParse(cleaned, out var parsedValue) ? parsedValue : 0.0;
}

static string FormatDollarAmount(double value)
{
    return string.Create(
        System.Globalization.CultureInfo.InvariantCulture,
        $"${value:0,0.00}");
}

static Regex DollarAmountRegex()
{
    return new Regex(@"\$(\d{1,3}(?:,\d{3})*\.\d{2})", RegexOptions.Compiled);
}

static string ApplyReplacements(
    string input,
    IReadOnlyList<ReplacementRule> replacements,
    out int replacementCount)
{
    var output = input;
    replacementCount = 0;

    foreach (var replacement in replacements)
    {
        if (string.IsNullOrEmpty(replacement.Original) || replacement.Original == replacement.Replacement)
        {
            continue;
        }

        var occurrences = CountOccurrences(output, replacement.Original);
        if (occurrences == 0)
        {
            continue;
        }

        output = output.Replace(replacement.Original, replacement.Replacement, StringComparison.Ordinal);
        replacementCount += occurrences;
    }

    return output;
}

static int CountOccurrences(string input, string value)
{
    if (string.IsNullOrEmpty(value))
    {
        return 0;
    }

    var count = 0;
    var index = 0;
    while ((index = input.IndexOf(value, index, StringComparison.Ordinal)) >= 0)
    {
        count++;
        index += value.Length;
    }

    return count;
}

static bool TryLoadXml(ZipArchiveEntry entry, out XDocument? document)
{
    try
    {
        using var stream = entry.Open();
        document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
        return true;
    }
    catch
    {
        document = null;
        return false;
    }
}

static void SaveXml(ZipArchiveEntry entry, XDocument document)
{
    using var writableStream = entry.Open();
    writableStream.SetLength(0);
    document.Save(writableStream);
}

static string NextValue(string[] args, ref int index, string optionName)
{
    if (index + 1 >= args.Length)
    {
        throw new ArgumentException($"Missing value for {optionName}.");
    }

    index++;
    return args[index];
}

static int Fail(string message)
{
    Console.Error.WriteLine(message);
    PrintUsage();
    return 1;
}

static void PrintUsage()
{
    Console.Error.WriteLine(
        "Usage:\n" +
        "  MatrixOpenXmlTextboxTool convert-docx-to-pdf --docx <path> --output <pdf-path>\n" +
        "  MatrixOpenXmlTextboxTool rewrite-matrix-amounts --docx <path>\n" +
        "  MatrixOpenXmlTextboxTool rewrite-textboxes --docx <path> --replacements <json-path>");
}

internal sealed class OpenXmlPartDocument
{
    public OpenXmlPartDocument(ZipArchiveEntry entry, XDocument document)
    {
        Entry = entry;
        Document = document;
    }

    public ZipArchiveEntry Entry { get; }
    public XDocument Document { get; }
    public bool Changed { get; set; }
}

internal sealed class ReplacementPayload
{
    public List<ReplacementRule> Replacements { get; init; } = new();
}

internal sealed class ReplacementRule
{
    public string Original { get; init; } = string.Empty;
    public string Replacement { get; init; } = string.Empty;
}
