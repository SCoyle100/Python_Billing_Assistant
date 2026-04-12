using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;

if (args.Length == 0)
{
    PrintUsage();
    return 1;
}

return args[0] switch
{
    "rewrite-textboxes" => RewriteTextboxes(args.Skip(1).ToArray()),
    _ => Fail($"Unknown command '{args[0]}'.")
};

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
        payload = JsonSerializer.Deserialize<ReplacementPayload>(File.ReadAllText(replacementsPath));
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
    return 0;
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
        "Usage: MatrixOpenXmlTextboxTool rewrite-textboxes --docx <path> --replacements <json-path>");
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
