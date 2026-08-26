using System.Text.Json;
using ClosedXML.Excel;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DrawingQC.Web;

/// <summary>One file that was added to a platform's consolidation, stamped with the day it arrived.</summary>
public sealed class ConsEntry
{
    public string Date { get; set; } = "";   // dd-MM-yyyy the file was added
    public string Name { get; set; } = "";    // original file name
    public int Count { get; set; }            // rows appended (Excel) or pages appended (PDF)
}

public sealed class CatManifest
{
    public int ExcelRev { get; set; }         // revisions shared (Rev 1,2,3…N) of the consolidated Excel
    public int PdfRev { get; set; }           // …and of the consolidated PDF
    public string LastExcelDate { get; set; } = ""; // last day a banner was written, so the date shows once per day
    public List<ConsEntry> ExcelEntries { get; set; } = new();
    public List<ConsEntry> PdfEntries { get; set; } = new();
}

/// <summary>
/// S2NERGY "ConsList": per platform (RP5S, WHP13N, …) and per category (Internal / External),
/// users add Excel + PDF files on a daily basis. Each Excel is appended, row-by-row and datewise,
/// into one running consolidated Excel (typical supports); each PDF's pages are appended into one
/// running consolidated PDF (unique supports). The file keeps growing — every download is the next
/// revision (Rev N) that can be shared with the client mid-month or at month-end.
/// </summary>
public static class ConsList
{
    private static readonly object Gate = new();
    public static readonly string[] Categories = { "Internal", "External" };
    private static readonly string[] DefaultProjects = { "RP5S", "WHP13N", "RP6N", "RP9S", "RP5N", "RP3S" };
    private static readonly JsonSerializerOptions JsonOpts = new() { WriteIndented = true };

    private static string Root()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        if (string.IsNullOrWhiteSpace(appData)) appData = AppContext.BaseDirectory;
        var dir = Path.Combine(appData, "SupportAutomation", "ConsList");
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static string ProjectsFile() => Path.Combine(Root(), "projects.json");
    private static string Safe(string s) => string.Concat((s ?? "").Split(Path.GetInvalidFileNameChars())).Trim();

    private static string CatDir(string platform, string category)
    {
        if (!Categories.Contains(category, StringComparer.OrdinalIgnoreCase)) category = "Internal";
        var dir = Path.Combine(Root(), Safe(platform), Safe(category));
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static string XlsxPath(string p, string c) => Path.Combine(CatDir(p, c), "consolidated.xlsx");
    private static string PdfPath(string p, string c) => Path.Combine(CatDir(p, c), "consolidated.pdf");
    private static string ManifestPath(string p, string c) => Path.Combine(CatDir(p, c), "manifest.json");

    // ---------- projects ----------

    public static List<string> Projects()
    {
        lock (Gate)
        {
            var f = ProjectsFile();
            if (!File.Exists(f))
            {
                var seed = DefaultProjects.ToList();
                File.WriteAllText(f, JsonSerializer.Serialize(seed, JsonOpts));
                return seed;
            }
            try { return JsonSerializer.Deserialize<List<string>>(File.ReadAllText(f)) ?? DefaultProjects.ToList(); }
            catch { return DefaultProjects.ToList(); }
        }
    }

    public static (bool ok, string err) AddProject(string name)
    {
        name = (name ?? "").Trim();
        if (name.Length == 0) return (false, "Enter a project / platform number.");
        lock (Gate)
        {
            var list = Projects();
            if (list.Any(p => p.Equals(name, StringComparison.OrdinalIgnoreCase)))
                return (false, "That project already exists.");
            list.Add(name);
            File.WriteAllText(ProjectsFile(), JsonSerializer.Serialize(list, JsonOpts));
            return (true, "");
        }
    }

    // Remove a project from the list and delete its stored consolidation (both categories).
    public static (bool ok, string err) RemoveProject(string name)
    {
        name = (name ?? "").Trim();
        if (name.Length == 0) return (false, "No project selected.");
        lock (Gate)
        {
            var list = Projects();
            var match = list.FirstOrDefault(p => p.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (match == null) return (false, "That project does not exist.");

            list.Remove(match);
            File.WriteAllText(ProjectsFile(), JsonSerializer.Serialize(list, JsonOpts));

            try
            {
                var dir = Path.Combine(Root(), Safe(match));
                if (Directory.Exists(dir)) Directory.Delete(dir, recursive: true);
            }
            catch { /* list is already updated; leftover files are harmless */ }
            return (true, "");
        }
    }

    // ---------- manifest ----------

    private static CatManifest LoadManifest(string p, string c)
    {
        var f = ManifestPath(p, c);
        if (!File.Exists(f)) return new CatManifest();
        try { return JsonSerializer.Deserialize<CatManifest>(File.ReadAllText(f)) ?? new CatManifest(); }
        catch { return new CatManifest(); }
    }

    private static void SaveManifest(string p, string c, CatManifest m) =>
        File.WriteAllText(ManifestPath(p, c), JsonSerializer.Serialize(m, JsonOpts));

    // ---------- merge ----------

    // Append an uploaded Excel's rows (verbatim, no extra column) onto the consolidated workbook.
    // When writeBanner is set, a single date row is written first so the date shows once per day.
    // Returns the number of content rows appended.
    private static int AppendExcel(string platform, string category, string sourcePath, string dateStr, bool writeBanner)
    {
        using var src = new XLWorkbook(sourcePath);
        var sws = src.Worksheets
            .Select(s => (s, rows: s.RangeUsed()?.RowCount() ?? 0))
            .OrderByDescending(x => x.rows).Select(x => x.s).FirstOrDefault();
        var used = sws?.RangeUsed();
        if (sws == null || used == null) return 0;

        int firstRow = used.FirstRow().RowNumber(), lastRow = used.LastRow().RowNumber();
        int firstCol = used.FirstColumn().ColumnNumber(), lastCol = used.LastColumn().ColumnNumber();

        var xlsx = XlsxPath(platform, category);
        XLWorkbook wb;
        IXLWorksheet ws;
        if (File.Exists(xlsx))
        {
            wb = new XLWorkbook(xlsx);
            if (!wb.TryGetWorksheet("Consolidated", out ws!)) ws = wb.Worksheets.First();
        }
        else
        {
            wb = new XLWorkbook();
            ws = wb.AddWorksheet("Consolidated");
        }

        int destRow = (ws.LastRowUsed()?.RowNumber() ?? 0) + 1;

        // Date shown once for the day, as a banner row above that day's data.
        if (writeBanner)
        {
            var b = ws.Cell(destRow, 1);
            b.Value = dateStr;
            b.Style.Font.Bold = true;
            b.Style.Fill.BackgroundColor = XLColor.FromArgb(0xDB, 0xE4, 0xF0);
            destRow++;
        }

        int added = 0;
        for (int r = firstRow; r <= lastRow; r++)
        {
            bool empty = true;
            for (int c = firstCol; c <= lastCol; c++)
                if (!sws.Cell(r, c).IsEmpty()) { empty = false; break; }
            if (empty) continue;

            int dc = 1;
            for (int c = firstCol; c <= lastCol; c++) ws.Cell(destRow, dc++).Value = sws.Cell(r, c).Value;
            destRow++; added++;
        }

        ws.Columns().AdjustToContents();
        wb.SaveAs(xlsx);
        wb.Dispose();
        return added;
    }

    // Append every page of an uploaded PDF onto the running consolidated PDF. Returns pages added.
    private static int AppendPdf(string platform, string category, string sourcePath)
    {
        var pdfPath = PdfPath(platform, category);
        PdfDocument outDoc = File.Exists(pdfPath)
            ? PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify)
            : new PdfDocument();
        int pages = 0;
        using (var s = PdfReader.Open(sourcePath, PdfDocumentOpenMode.Import))
            for (int i = 0; i < s.PageCount; i++) { outDoc.AddPage(s.Pages[i]); pages++; }
        outDoc.Save(pdfPath);
        outDoc.Dispose();
        return pages;
    }

    /// <summary>
    /// Add a batch of uploaded files (already buffered to temp paths). Each .xlsx/.xls appends to the
    /// consolidated Excel; each .pdf appends to the consolidated PDF. Returns per-batch totals + any errors.
    /// </summary>
    public static (int excelRows, int pdfPages, int excelFiles, int pdfFiles, List<string> errors) Add(
        string platform, string category, IEnumerable<(string name, string tempPath)> files)
    {
        lock (Gate)
        {
            var m = LoadManifest(platform, category);
            string date = DateTime.Now.ToString("dd-MM-yyyy");
            int rows = 0, pages = 0, ef = 0, pf = 0;
            var errors = new List<string>();
            bool bannerPending = !string.Equals(m.LastExcelDate, date, StringComparison.Ordinal);

            foreach (var (name, temp) in files)
            {
                var ext = Path.GetExtension(name).ToLowerInvariant();
                try
                {
                    if (ext is ".xlsx" or ".xls")
                    {
                        int n = AppendExcel(platform, category, temp, date, bannerPending);
                        bannerPending = false;        // date banner is written once per day
                        m.LastExcelDate = date;
                        m.ExcelEntries.Add(new ConsEntry { Date = date, Name = name, Count = n });
                        rows += n; ef++;
                    }
                    else if (ext == ".pdf")
                    {
                        int n = AppendPdf(platform, category, temp);
                        m.PdfEntries.Add(new ConsEntry { Date = date, Name = name, Count = n });
                        pages += n; pf++;
                    }
                    else errors.Add($"{name}: only .xlsx/.xls and .pdf are supported.");
                }
                catch (Exception ex) { errors.Add($"{name}: {(string.IsNullOrWhiteSpace(ex.Message) ? ex.GetType().Name : ex.Message)}"); }
            }

            SaveManifest(platform, category, m);
            return (rows, pages, ef, pf, errors);
        }
    }

    /// <summary>Full state for the UI: projects + per platform/category totals and the datewise add log.</summary>
    public static object State()
    {
        lock (Gate)
        {
            var projects = Projects();
            var data = new Dictionary<string, object>();
            foreach (var p in projects)
            {
                var cats = new Dictionary<string, object>();
                foreach (var c in Categories)
                {
                    var m = LoadManifest(p, c);
                    var log = m.ExcelEntries.Select(e => new { e.Date, e.Name, e.Count, type = "Excel" })
                        .Concat(m.PdfEntries.Select(e => new { e.Date, e.Name, e.Count, type = "PDF" }))
                        .OrderBy(e => DateTime.TryParseExact(e.Date, "dd-MM-yyyy", null,
                            System.Globalization.DateTimeStyles.None, out var d) ? d : DateTime.MaxValue)
                        .ToList();
                    cats[c] = new
                    {
                        excelRev = m.ExcelRev,
                        pdfRev = m.PdfRev,
                        excelRows = m.ExcelEntries.Sum(e => e.Count),
                        pdfPages = m.PdfEntries.Sum(e => e.Count),
                        excelFiles = m.ExcelEntries.Count,
                        pdfFiles = m.PdfEntries.Count,
                        hasExcel = File.Exists(XlsxPath(p, c)),
                        hasPdf = File.Exists(PdfPath(p, c)),
                        log,
                    };
                }
                data[p] = cats;
            }
            return new { projects, categories = Categories, data };
        }
    }

    /// <summary>Read the consolidated file for download, bumping its revision (each share = the next Rev N).</summary>
    public static (byte[]? bytes, string fileName, int rev, string? err) Download(string platform, string category, string type)
    {
        lock (Gate)
        {
            bool excel = type.Equals("excel", StringComparison.OrdinalIgnoreCase);
            var path = excel ? XlsxPath(platform, category) : PdfPath(platform, category);
            if (!File.Exists(path))
                return (null, "", 0, excel ? "No Excel has been added for this platform yet." : "No PDF has been added for this platform yet.");

            var m = LoadManifest(platform, category);
            int rev = excel ? ++m.ExcelRev : ++m.PdfRev;
            SaveManifest(platform, category, m);

            var bytes = File.ReadAllBytes(path);
            string ext = excel ? "xlsx" : "pdf";
            string fileName = $"{Safe(platform)}_{Safe(category)}_Consolidated_Rev{rev}_{DateTime.Now:yyyy-MM-dd}.{ext}";
            return (bytes, fileName, rev, null);
        }
    }
}
