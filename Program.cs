using System.IO.Compression;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using UglyToad.PdfPig;

namespace DrawingQC;

internal static class Program
{
    // Matches drawing numbers like: 265-S2N-ACD-5003, 265-S2N-BCD-4308, 265-S2N-L2-2231
    // Format: 3 digits - alphanumeric - alphanumeric - 4 digits.
    // Last segment is pinned to exactly 4 digits: PDF text extraction often glues
    // adjacent title-block text onto the number (e.g. "...-5003" -> "...-500330").
    private static readonly Regex DrawingNumberRegex =
        new(@"\d{3}-[A-Z0-9]{2,5}-[A-Z0-9]{1,5}-\d{4}", RegexOptions.Compiled | RegexOptions.IgnoreCase);

    private static int Main(string[] args)
    {
        string zipPath = args.Length > 0
            ? args[0]
            : @"C:\Users\pushpat\Downloads\WHP13N ALL PDF.zip";

        string outputPath = args.Length > 1
            ? args[1]
            : Path.Combine(
                Path.GetDirectoryName(zipPath) ?? ".",
                Path.GetFileNameWithoutExtension(zipPath) + " - QC Report.xlsx");

        if (!File.Exists(zipPath))
        {
            Console.Error.WriteLine($"ERROR: Zip file not found: {zipPath}");
            return 1;
        }

        Console.WriteLine($"Reading zip : {zipPath}");
        Console.WriteLine($"Report      : {outputPath}");
        Console.WriteLine();

        var results = new List<PdfResult>();

        using (var archive = ZipFile.OpenRead(zipPath))
        {
            var pdfEntries = archive.Entries
                .Where(e => e.Name.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
                .OrderBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            int index = 0;
            foreach (var entry in pdfEntries)
            {
                index++;
                Console.Write($"[{index}/{pdfEntries.Count}] {entry.Name} ... ");
                var result = ProcessPdf(entry);
                results.Add(result);
                Console.WriteLine(result.MatchStatus);
            }
        }

        // ---- Duplicate detection across the whole set ----
        // A drawing name is a duplicate if the same number appears in more than one place
        // (in another PDF, or twice within the same PDF).
        var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var r in results)
            foreach (var key in r.DrawingKeys)
                if (!string.IsNullOrWhiteSpace(key))
                    counts[key] = counts.GetValueOrDefault(key) + 1;

        foreach (var r in results)
            r.IsDuplicate = r.DrawingKeys.Any(k => !string.IsNullOrWhiteSpace(k) && counts.GetValueOrDefault(k) > 1);

        WriteReport(results, outputPath);

        int matched = results.Count(r => r.FinalStatus == "Matched");
        int unmatched = results.Count(r => r.FinalStatus == "Unmatched");
        int duplicate = results.Count(r => r.FinalStatus == "Duplicate");
        Console.WriteLine();
        Console.WriteLine($"Done. Total: {results.Count}  |  Matched: {matched}  |  Unmatched: {unmatched}  |  Duplicate: {duplicate}");
        Console.WriteLine($"Excel report written to: {outputPath}");

        if (Environment.GetEnvironmentVariable("QC_DEBUG") == "1")
        {
            Console.WriteLine("\n--- Duplicate / Unmatched rows ---");
            foreach (var r in results.Where(x => x.FinalStatus != "Matched"))
                Console.WriteLine($"[{r.FinalStatus}] {r.FileName}  ->  1st={r.Drawing1} | 2nd={r.Drawing2}");
        }
        return 0;
    }

    private static PdfResult ProcessPdf(ZipArchiveEntry entry)
    {
        string fileName = entry.Name;
        string nameNoExt = Path.GetFileNameWithoutExtension(fileName);

        // Split the file name into the individual drawings on "&".
        var segments = nameNoExt.Split('&', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        var drawingDisplays = new List<string>();
        var drawingKeys = new List<string>();
        foreach (var seg in segments)
        {
            var m = DrawingNumberRegex.Match(seg);
            string display = m.Success ? m.Value.ToUpperInvariant() : CleanName(seg);
            drawingDisplays.Add(display);
            drawingKeys.Add(display);
        }

        var result = new PdfResult
        {
            FileName = fileName,
            Drawing1 = drawingDisplays.ElementAtOrDefault(0) ?? "",
            Drawing2 = drawingDisplays.ElementAtOrDefault(1) ?? "",
            DrawingKeys = drawingKeys,
            NameNumbers = ExtractNumbers(nameNoExt),
        };

        try
        {
            using var ms = new MemoryStream();
            using (var s = entry.Open())
                s.CopyTo(ms);
            ms.Position = 0;

            var contentNumbers = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var doc = PdfDocument.Open(ms))
                foreach (var page in doc.GetPages())
                    foreach (var num in ExtractNumbers(page.Text))
                        contentNumbers.Add(num);

            result.ContentNumbers = contentNumbers;
        }
        catch (Exception ex)
        {
            result.Error = ex.Message;
        }

        return result;
    }

    private static string CleanName(string s)
    {
        // Strip leading "Copy of " (possibly repeated) and normalise.
        var t = s.Trim();
        while (t.StartsWith("Copy of ", StringComparison.OrdinalIgnoreCase))
            t = t.Substring("Copy of ".Length).Trim();
        return t.ToUpperInvariant();
    }

    private static SortedSet<string> ExtractNumbers(string text)
    {
        var set = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrEmpty(text))
            return set;
        foreach (Match m in DrawingNumberRegex.Matches(text))
            set.Add(m.Value.ToUpperInvariant());
        return set;
    }

    private static void WriteReport(List<PdfResult> results, string outputPath)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Drawing QC");

        string[] headers = { "Sr.No", "Support PDF", "1st drawing name", "2nd drawing name", "Status" };
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(1, c + 1);
            cell.Value = headers[c];
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Fill.BackgroundColor = XLColor.FromArgb(0x37, 0x47, 0x5A);
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        }

        var matchFill = XLColor.FromArgb(0xC6, 0xEF, 0xCE);    // green
        var matchFont = XLColor.FromArgb(0x00, 0x61, 0x00);
        var unmatchFill = XLColor.FromArgb(0xFF, 0xC7, 0xCE);  // red
        var unmatchFont = XLColor.FromArgb(0x9C, 0x00, 0x06);
        var dupFill = XLColor.FromArgb(0xFF, 0xEB, 0x9C);      // amber/yellow
        var dupFont = XLColor.FromArgb(0x9C, 0x63, 0x00);

        int row = 2;
        foreach (var r in results)
        {
            ws.Cell(row, 1).Value = row - 1;
            ws.Cell(row, 2).Value = r.FileName;
            ws.Cell(row, 3).Value = string.IsNullOrWhiteSpace(r.Drawing1) ? "-" : r.Drawing1;
            ws.Cell(row, 4).Value = string.IsNullOrWhiteSpace(r.Drawing2) ? "-" : r.Drawing2;
            ws.Cell(row, 5).Value = r.FinalStatus;

            XLColor fill, font;
            switch (r.FinalStatus)
            {
                case "Matched": fill = matchFill; font = matchFont; break;
                case "Duplicate": fill = dupFill; font = dupFont; break;
                default: fill = unmatchFill; font = unmatchFont; break;
            }

            var range = ws.Range(row, 1, row, headers.Length);
            range.Style.Fill.BackgroundColor = fill;
            range.Style.Font.FontColor = font;
            row++;
        }

        ws.SheetView.FreezeRows(1);
        ws.Range(1, 1, Math.Max(row - 1, 1), headers.Length).SetAutoFilter();
        ws.Column(1).Width = 8;
        ws.Column(2).Width = 55;
        ws.Column(3).Width = 24;
        ws.Column(4).Width = 24;
        ws.Column(5).Width = 14;

        // Summary sheet
        var sum = wb.AddWorksheet("Summary");
        sum.Cell(1, 1).Value = "Drawing QC Summary";
        sum.Cell(1, 1).Style.Font.Bold = true;
        sum.Cell(1, 1).Style.Font.FontSize = 14;
        sum.Cell(3, 1).Value = "Total PDFs";
        sum.Cell(3, 2).Value = results.Count;
        sum.Cell(4, 1).Value = "Matched";
        sum.Cell(4, 2).Value = results.Count(x => x.FinalStatus == "Matched");
        sum.Cell(4, 1).Style.Fill.BackgroundColor = matchFill;
        sum.Cell(4, 2).Style.Fill.BackgroundColor = matchFill;
        sum.Cell(5, 1).Value = "Unmatched";
        sum.Cell(5, 2).Value = results.Count(x => x.FinalStatus == "Unmatched");
        sum.Cell(5, 1).Style.Fill.BackgroundColor = unmatchFill;
        sum.Cell(5, 2).Style.Fill.BackgroundColor = unmatchFill;
        sum.Cell(6, 1).Value = "Duplicate";
        sum.Cell(6, 2).Value = results.Count(x => x.FinalStatus == "Duplicate");
        sum.Cell(6, 1).Style.Fill.BackgroundColor = dupFill;
        sum.Cell(6, 2).Style.Fill.BackgroundColor = dupFill;
        sum.Column(1).Width = 24;
        sum.Column(2).Width = 12;

        wb.SaveAs(outputPath);
    }
}

internal sealed class PdfResult
{
    public string FileName { get; set; } = "";
    public string Drawing1 { get; set; } = "";
    public string Drawing2 { get; set; } = "";
    public List<string> DrawingKeys { get; set; } = new();
    public SortedSet<string> NameNumbers { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public SortedSet<string> ContentNumbers { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public string? Error { get; set; }
    public bool IsDuplicate { get; set; }

    // Matched = every drawing number in the file name is found inside the PDF, and vice versa.
    public bool IsMatch =>
        Error == null &&
        NameNumbers.Count > 0 &&
        NameNumbers.SetEquals(ContentNumbers);

    public string MatchStatus => IsMatch ? "Matched" : "Unmatched";

    // Duplicate takes priority so repeated drawings surface even if they otherwise match.
    public string FinalStatus =>
        IsDuplicate ? "Duplicate" : (IsMatch ? "Matched" : "Unmatched");
}
