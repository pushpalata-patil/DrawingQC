using System.Text.Json;
using ClosedXML.Excel;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DrawingQC.Web;

/// <summary>One file that was added to a platform's consolidation, stamped with the day it arrived.</summary>
public sealed class ConsEntry
{
    public string Id { get; set; } = "";      // stable id; the source file is stored as <Id><Ext>
    public string Ext { get; set; } = "";      // ".xlsx" / ".xls" / ".pdf"
    public string Date { get; set; } = "";     // dd-MM-yyyy the file was added
    public string Name { get; set; } = "";     // original file name
    public int Count { get; set; }             // rows (Excel) or pages (PDF) this file contributes
}

public sealed class CatManifest
{
    public int ExcelRev { get; set; }
    public int PdfRev { get; set; }
    public string LastExcelDate { get; set; } = "";
    public List<ConsEntry> ExcelEntries { get; set; } = new();
    public List<ConsEntry> PdfEntries { get; set; } = new();
}

/// <summary>
/// S2NERGY "ConsList": per platform (RP5S, WHP13N, …) and per category (Internal / External),
/// users add MTO Excel + Unique PDF files daily. Every uploaded file is kept as a source; the
/// consolidated Excel (rows, datewise banner) and consolidated PDF (pages) are rebuilt from those
/// sources, so individual files can be deleted or replaced. Downloads are Internal / External /
/// Combined, each bumping its own revision (Rev N).
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
    private static string SourcesDir(string p, string c) { var d = Path.Combine(CatDir(p, c), "sources"); Directory.CreateDirectory(d); return d; }
    private static string SourcePath(string p, string c, ConsEntry e) => Path.Combine(SourcesDir(p, c), e.Id + e.Ext);
    // ---------- blob mirror (hosted mode) ----------
    // With PostgreSQL configured the container's disk is only a cache: every managed file (source
    // uploads, consolidated Excel/PDF) is mirrored into the blobs table on write and pulled back on
    // demand, so a redeploy / restart / spin-down on an ephemeral filesystem loses nothing.
    private static string BlobKey(string path) => Path.GetRelativePath(Root(), path).Replace('\\', '/');

    /// <summary>True if the file is on disk, or could be restored from the database.</summary>
    private static bool Have(string path)
    {
        if (File.Exists(path)) return true;
        if (!Db.Enabled) return false;
        try
        {
            var data = Db.GetBlob(BlobKey(path));
            if (data == null) return false;
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllBytes(path, data);
            return true;
        }
        catch (Exception ex) { Console.Error.WriteLine("[ConsList] blob fetch failed: " + ex.Message); return false; }
    }

    /// <summary>Mirror a freshly written file into the database.</summary>
    private static void Saved(string path)
    {
        if (!Db.Enabled || !File.Exists(path)) return;
        try { Db.PutBlob(BlobKey(path), File.ReadAllBytes(path)); }
        catch (Exception ex) { Console.Error.WriteLine("[ConsList] blob store failed: " + ex.Message); }
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
        if (Db.Enabled) { try { Db.DeleteBlob(BlobKey(path)); } catch { } }
    }

    // ---------- projects ----------

    public static List<string> Projects()
    {
        if (Db.Enabled)
        {
            var l = Db.LoadProjects();
            if (l.Count == 0) { l = DefaultProjects.ToList(); Db.SaveProjects(l); }
            return l;
        }
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

    private static void PersistProjects(List<string> list)
    {
        if (Db.Enabled) { Db.SaveProjects(list); return; }
        File.WriteAllText(ProjectsFile(), JsonSerializer.Serialize(list, JsonOpts));
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
            PersistProjects(list);
            return (true, "");
        }
    }

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
            PersistProjects(list);
            if (Db.Enabled) Db.DeletePlatformData(match);            // remove its file/category rows
            if (Db.Enabled) { try { Db.DeleteBlobsWithPrefix(Safe(match) + "/"); } catch { } }
            try { var dir = Path.Combine(Root(), Safe(match)); if (Directory.Exists(dir)) Directory.Delete(dir, true); }
            catch { }
            return (true, "");
        }
    }

    // ---------- manifest ----------

    private static CatManifest LoadManifest(string p, string c)
    {
        if (Db.Enabled) return Db.LoadManifest(p, c);
        var f = ManifestPath(p, c);
        if (!File.Exists(f)) return new CatManifest();
        try { return JsonSerializer.Deserialize<CatManifest>(File.ReadAllText(f)) ?? new CatManifest(); }
        catch { return new CatManifest(); }
    }

    private static void SaveManifest(string p, string c, CatManifest m)
    {
        if (Db.Enabled) { Db.SaveManifest(p, c, m); return; }
        File.WriteAllText(ManifestPath(p, c), JsonSerializer.Serialize(m, JsonOpts));
    }

    // Assign a stable id + extension to any legacy entry that lacks them. Returns true if changed.
    private static bool EnsureIds(CatManifest m)
    {
        bool ch = false;
        foreach (var e in m.ExcelEntries)
        {
            if (string.IsNullOrEmpty(e.Id)) { e.Id = Guid.NewGuid().ToString("N"); ch = true; }
            if (string.IsNullOrEmpty(e.Ext)) { var x = Path.GetExtension(e.Name).ToLowerInvariant(); e.Ext = (x == ".xls" || x == ".xlsx") ? x : ".xlsx"; ch = true; }
        }
        foreach (var e in m.PdfEntries)
        {
            if (string.IsNullOrEmpty(e.Id)) { e.Id = Guid.NewGuid().ToString("N"); ch = true; }
            if (string.IsNullOrEmpty(e.Ext)) { e.Ext = ".pdf"; ch = true; }
        }
        return ch;
    }

    // ---------- source backfill (one-time, for data added before sources were kept) ----------

    // Split the existing consolidated Excel back into per-entry source files (data rows only, no banner).
    private static void BackfillExcel(string p, string c, CatManifest m)
    {
        if (!Have(XlsxPath(p, c))) return;
        using var wb = new XLWorkbook(XlsxPath(p, c));
        if (!wb.TryGetWorksheet("Consolidated", out var ws)) ws = wb.Worksheets.First();
        var used = ws.RangeUsed(); if (used == null) return;
        int maxCol = used.LastColumn().ColumnNumber();
        int r = 1; string? lastDate = null;
        foreach (var e in m.ExcelEntries)
        {
            if (e.Date != lastDate) { r++; lastDate = e.Date; }   // skip the once-per-day banner row
            var sp = SourcePath(p, c, e);
            if (!Have(sp))
            {
                using var ow = new XLWorkbook(); var os = ow.AddWorksheet("Sheet1");
                for (int k = 0; k < e.Count; k++)
                    for (int col = 1; col <= maxCol; col++) os.Cell(k + 1, col).Value = ws.Cell(r + k, col).Value;
                ow.SaveAs(sp);
                Saved(sp);
            }
            r += e.Count;
        }
    }

    // Split the existing consolidated PDF back into per-entry source files by page count.
    private static void BackfillPdf(string p, string c, CatManifest m)
    {
        if (!Have(PdfPath(p, c))) return;
        var bytes = File.ReadAllBytes(PdfPath(p, c));
        int page = 0;
        foreach (var e in m.PdfEntries)
        {
            var sp = SourcePath(p, c, e);
            if (!Have(sp))
            {
                using var src = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
                using var od = new PdfDocument();
                for (int k = 0; k < e.Count && page + k < src.PageCount; k++) od.AddPage(src.Pages[page + k]);
                od.Save(sp);
                Saved(sp);
            }
            page += e.Count;
        }
    }

    private static void EnsureSources(string p, string c, CatManifest m)
    {
        EnsureIds(m);
        if (m.ExcelEntries.Any(e => !Have(SourcePath(p, c, e)))) BackfillExcel(p, c, m);
        if (m.PdfEntries.Any(e => !Have(SourcePath(p, c, e)))) BackfillPdf(p, c, m);
    }

    // ---------- build ----------

    // First row that holds data: the first row whose leading (SL NO) column is a positive integer.
    // Everything above it is the template's header block (logo, title, subtitle, coloured headers).
    private static int DataStartRow(IXLWorksheet ws)
    {
        var u = ws.RangeUsed();
        if (u == null) return 1;
        int fr = u.FirstRow().RowNumber(), lr = u.LastRow().RowNumber(), fc = u.FirstColumn().ColumnNumber();
        for (int r = fr; r <= lr; r++)
        {
            var s = ws.Cell(r, fc).GetString().Trim();
            if (int.TryParse(s, out var n) && n >= 1) return r;
        }
        return fr; // no serial column found — treat everything as data
    }


    // Rebuild the consolidated Excel by reproducing every uploaded file VERBATIM, stacked one below
    // the next. Each block keeps that file's own layout exactly as uploaded — logo, title, coloured
    // headers, its own columns and data — and is introduced by a row showing the date it was added.
    // Files with different column counts are each kept at their own width; nothing is merged or
    // aligned.
    private static void RebuildExcel(string p, string c, CatManifest m)
    {
        var xlsx = XlsxPath(p, c);
        if (m.ExcelEntries.Count == 0) { TryDelete(xlsx); m.LastExcelDate = ""; return; }

        using var outWb = new XLWorkbook();
        var outWs = outWb.AddWorksheet(c);   // sheet named "Internal" / "External"
        int outRow = 1;
        int picN = 0;
        bool any = false;

        foreach (var e in m.ExcelEntries)
        {
            var sp = SourcePath(p, c, e);
            if (!Have(sp)) { e.Count = 0; continue; }
            using var wb = new XLWorkbook(sp);
            var ws = wb.Worksheets.OrderByDescending(s => s.RangeUsed()?.RowCount() ?? 0).FirstOrDefault();
            var u = ws?.RangeUsed();
            if (ws == null || u == null) { e.Count = 0; continue; }
            int lastR = u.LastRow().RowNumber(), lastC = u.LastColumn().ColumnNumber();
            any = true;

            int blockTop = outRow;
            // Copy every row (from 1, so the logo/title band comes too) verbatim: values + style + height.
            for (int r = 1; r <= lastR; r++)
            {
                try { outWs.Row(outRow).Height = ws.Row(r).Height; } catch { }
                for (int col = 1; col <= lastC; col++)
                {
                    var s = ws.Cell(r, col); var d = outWs.Cell(outRow, col);
                    d.Value = s.Value; d.Style = s.Style;
                }
                outRow++;
            }

            // Re-create the merged ranges, offset to this block's position.
            foreach (var mr in ws.MergedRanges)
            {
                var a = mr.RangeAddress;
                int fr = a.FirstAddress.RowNumber, fcc = a.FirstAddress.ColumnNumber;
                int lr = a.LastAddress.RowNumber, lcc = a.LastAddress.ColumnNumber;
                try { outWs.Range(blockTop + fr - 1, fcc, blockTop + lr - 1, lcc).Merge(); } catch { }
            }

            // Column widths: keep the widest seen for each column across all files.
            for (int col = 1; col <= lastC; col++)
            {
                double w = ws.Column(col).Width;
                if (w > outWs.Column(col).Width) outWs.Column(col).Width = w;
            }

            // Re-place each picture (logo) at this block's top, sized to fit the header (native
            // dimensions are the full-resolution image, so set an explicit display size).
            foreach (var pic in ws.Pictures)
            {
                try
                {
                    var img = pic.ImageStream; img.Position = 0;
                    double aspect = pic.Height > 0 ? (double)pic.Width / pic.Height : 3.25;
                    int h = 80, wpx = (int)(h * aspect);
                    int ar = pic.TopLeftCell?.Address.RowNumber ?? 1;
                    int ac = pic.TopLeftCell?.Address.ColumnNumber ?? 1;
                    outWs.AddPicture(img, pic.Format, $"logo{picN++}")
                         .MoveTo(outWs.Cell(blockTop + ar - 1, ac)).WithSize(wpx, h);
                }
                catch { }
            }

            e.Count = CountSupports(ws, u);
            outRow++; // blank gap before the next file's block
        }

        if (!any) { TryDelete(xlsx); m.LastExcelDate = ""; return; }
        outWb.SaveAs(xlsx);
        Saved(xlsx);
        m.LastExcelDate = m.ExcelEntries[^1].Date;
    }

    // Number of support rows in a sheet (rows with a value in the leading SL NO column, from the
    // first data row down) — used only for the on-screen count, not for the output.
    private static int CountSupports(IXLWorksheet ws, IXLRange used)
    {
        int fc = used.FirstColumn().ColumnNumber(), lr = used.LastRow().RowNumber();
        int ds = DataStartRow(ws), n = 0;
        for (int r = ds; r <= lr; r++) if (!ws.Cell(r, fc).IsEmpty()) n++;
        return n;
    }

    // Rebuild the consolidated PDF from all PDF sources (recomputing each entry's page count).
    private static void RebuildPdf(string p, string c, CatManifest m)
    {
        var pdfPath = PdfPath(p, c);
        if (m.PdfEntries.Count == 0) { TryDelete(pdfPath); return; }
        using var outDoc = new PdfDocument();
        foreach (var e in m.PdfEntries)
        {
            var sp = SourcePath(p, c, e);
            if (!Have(sp)) { e.Count = 0; continue; }
            using var s = PdfReader.Open(sp, PdfDocumentOpenMode.Import);
            int cnt = 0;
            for (int i = 0; i < s.PageCount; i++) { outDoc.AddPage(s.Pages[i]); cnt++; }
            e.Count = cnt;
        }
        outDoc.Save(pdfPath);
        Saved(pdfPath);
    }

    // ---------- add ----------

    /// <summary>
    /// Add a batch of uploaded files. Each file is saved as a source and appended to the consolidated
    /// Excel/PDF. Returns per-batch totals + any errors.
    /// </summary>
    public static (int excelRows, int pdfPages, int excelFiles, int pdfFiles, List<string> errors) Add(
        string platform, string category, IEnumerable<(string name, string tempPath)> files)
    {
        lock (Gate)
        {
            var m = LoadManifest(platform, category);
            EnsureIds(m);
            string date = DateTime.Now.ToString("dd-MM-yyyy");
            int rows = 0, pages = 0, ef = 0, pf = 0;
            var errors = new List<string>();
            var newExcel = new List<ConsEntry>();

            foreach (var (name, temp) in files)
            {
                var ext = Path.GetExtension(name).ToLowerInvariant();
                try
                {
                    if (ext is ".xlsx" or ".xls")
                    {
                        // Files are aligned by column header in RebuildExcel, so differing column
                        // sets (14 vs 18) merge cleanly — nothing is skipped for format.
                        var e = new ConsEntry { Id = Guid.NewGuid().ToString("N"), Ext = ext, Date = date, Name = name };
                        File.Copy(temp, SourcePath(platform, category, e), true);
                        Saved(SourcePath(platform, category, e));
                        m.ExcelEntries.Add(e); newExcel.Add(e); ef++;
                    }
                    else if (ext == ".pdf")
                    {
                        var e = new ConsEntry { Id = Guid.NewGuid().ToString("N"), Ext = ".pdf", Date = date, Name = name };
                        File.Copy(temp, SourcePath(platform, category, e), true);
                        Saved(SourcePath(platform, category, e));
                        int n = AppendPdf(platform, category, SourcePath(platform, category, e));
                        e.Count = n; m.PdfEntries.Add(e);
                        pages += n; pf++;
                    }
                    else errors.Add($"{name}: only .xlsx/.xls and .pdf are supported.");
                }
                catch (Exception ex) { errors.Add($"{name}: {(string.IsNullOrWhiteSpace(ex.Message) ? ex.GetType().Name : ex.Message)}"); }
            }

            // Rebuild the styled consolidated Excel from all sources so the template (logo, coloured
            // headers, support count) is preserved and each entry's row count is recomputed.
            if (newExcel.Count > 0) RebuildExcel(platform, category, m);
            rows = newExcel.Sum(e => e.Count);

            SaveManifest(platform, category, m);
            return (rows, pages, ef, pf, errors);
        }
    }

    // Incrementally append one PDF source onto the existing consolidated PDF (used by Add).
    private static int AppendPdf(string platform, string category, string sourcePath)
    {
        var pdfPath = PdfPath(platform, category);
        PdfDocument outDoc = Have(pdfPath) ? PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify) : new PdfDocument();
        int pages = 0;
        using (var s = PdfReader.Open(sourcePath, PdfDocumentOpenMode.Import))
            for (int i = 0; i < s.PageCount; i++) { outDoc.AddPage(s.Pages[i]); pages++; }
        outDoc.Save(pdfPath);
        Saved(pdfPath);
        outDoc.Dispose();
        return pages;
    }

    // ---------- delete / replace ----------

    /// <summary>Delete the given files (by id) across both categories, then rebuild what's left.</summary>
    public static (bool ok, string err, int deleted) DeleteFiles(string platform, IEnumerable<string> ids)
    {
        lock (Gate)
        {
            var idset = new HashSet<string>((ids ?? Enumerable.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s)));
            if (idset.Count == 0) return (false, "No files selected.", 0);

            int deleted = 0;
            foreach (var c in Categories)
            {
                var m = LoadManifest(platform, c);
                EnsureIds(m);
                bool touchedExcel = m.ExcelEntries.Any(e => idset.Contains(e.Id));
                bool touchedPdf = m.PdfEntries.Any(e => idset.Contains(e.Id));
                if (!touchedExcel && !touchedPdf) continue;

                EnsureSources(platform, c, m);   // make sure remaining files can be rebuilt
                foreach (var e in m.ExcelEntries.Where(e => idset.Contains(e.Id)).ToList()) { TryDelete(SourcePath(platform, c, e)); m.ExcelEntries.Remove(e); deleted++; }
                foreach (var e in m.PdfEntries.Where(e => idset.Contains(e.Id)).ToList()) { TryDelete(SourcePath(platform, c, e)); m.PdfEntries.Remove(e); deleted++; }
                if (touchedExcel) RebuildExcel(platform, c, m);
                if (touchedPdf) RebuildPdf(platform, c, m);
                SaveManifest(platform, c, m);
            }
            return deleted == 0 ? (false, "Selected files were not found.", 0) : (true, "", deleted);
        }
    }

    /// <summary>Replace one file (by id) with a newly uploaded file of the same kind, then rebuild.</summary>
    public static (bool ok, string err) ReplaceFile(string platform, string id, string name, string tempPath)
    {
        lock (Gate)
        {
            if (string.IsNullOrWhiteSpace(id)) return (false, "No file selected to replace.");
            var ext = Path.GetExtension(name).ToLowerInvariant();
            foreach (var c in Categories)
            {
                var m = LoadManifest(platform, c);
                EnsureIds(m);
                var xe = m.ExcelEntries.FirstOrDefault(e => e.Id == id);
                var pe = m.PdfEntries.FirstOrDefault(e => e.Id == id);
                if (xe == null && pe == null) continue;

                EnsureSources(platform, c, m);
                if (xe != null)
                {
                    if (ext is not (".xlsx" or ".xls")) return (false, "Replace an Excel file with an .xlsx / .xls file.");
                    TryDelete(SourcePath(platform, c, xe));
                    xe.Ext = ext; xe.Name = name;
                    File.Copy(tempPath, SourcePath(platform, c, xe), true);
                    Saved(SourcePath(platform, c, xe));
                    RebuildExcel(platform, c, m);
                }
                else
                {
                    if (ext != ".pdf") return (false, "Replace a PDF with a .pdf file.");
                    File.Copy(tempPath, SourcePath(platform, c, pe!), true);
                    Saved(SourcePath(platform, c, pe!));
                    pe!.Name = name;
                    RebuildPdf(platform, c, m);
                }
                SaveManifest(platform, c, m);
                return (true, "");
            }
            return (false, "That file was not found.");
        }
    }

    // ---------- state ----------

    /// <summary>Full state for the UI: projects + per platform/category totals and the datewise file log.</summary>
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
                    if (EnsureIds(m)) SaveManifest(p, c, m);   // so the UI can reference each file
                    var log = m.ExcelEntries.Select(e => new { e.Id, e.Date, e.Name, e.Count, type = "Excel" })
                        .Concat(m.PdfEntries.Select(e => new { e.Id, e.Date, e.Name, e.Count, type = "PDF" }))
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
                        hasExcel = Have(XlsxPath(p, c)),
                        hasPdf = Have(PdfPath(p, c)),
                        log,
                    };
                }
                data[p] = cats;
            }
            return new { projects, categories = Categories, data };
        }
    }

    // ---------- download ----------

    private sealed class CombinedRev { public int ExcelRev { get; set; } public int PdfRev { get; set; } }
    private static string CombinedFile(string p) => Path.Combine(Root(), Safe(p), "_combined.json");

    private static int BumpCombinedRev(string platform, bool excel)
    {
        if (Db.Enabled)
        {
            var (e, p) = Db.LoadCombinedRev(platform);
            if (excel) e++; else p++;
            Db.SaveCombinedRev(platform, e, p);
            return excel ? e : p;
        }
        Directory.CreateDirectory(Path.Combine(Root(), Safe(platform)));
        var f = CombinedFile(platform);
        CombinedRev cr;
        try { cr = File.Exists(f) ? (JsonSerializer.Deserialize<CombinedRev>(File.ReadAllText(f)) ?? new()) : new(); }
        catch { cr = new(); }
        int rev = excel ? ++cr.ExcelRev : ++cr.PdfRev;
        File.WriteAllText(f, JsonSerializer.Serialize(cr, JsonOpts));
        return rev;
    }

    private static (byte[]? bytes, string? err) BuildCombinedExcel(string platform)
    {
        var iPath = XlsxPath(platform, "Internal");
        var ePath = XlsxPath(platform, "External");
        bool hasI = Have(iPath), hasE = Have(ePath);
        if (!hasI && !hasE) return (null, "No Excel has been added for Internal or External yet.");
        using var outWb = new XLWorkbook();
        if (hasI) { using var w = new XLWorkbook(iPath); w.Worksheets.First().CopyTo(outWb, "Internal"); }
        if (hasE) { using var w = new XLWorkbook(ePath); w.Worksheets.First().CopyTo(outWb, "External"); }
        using var ms = new MemoryStream();
        outWb.SaveAs(ms);
        return (ms.ToArray(), null);
    }

    private static (byte[]? bytes, string? err) BuildCombinedPdf(string platform)
    {
        var iPath = PdfPath(platform, "Internal");
        var ePath = PdfPath(platform, "External");
        bool hasI = Have(iPath), hasE = Have(ePath);
        if (!hasI && !hasE) return (null, "No PDF has been added for Internal or External yet.");
        using var outDoc = new PdfDocument();
        foreach (var (has, path) in new[] { (hasI, iPath), (hasE, ePath) })
        {
            if (!has) continue;
            using var s = PdfReader.Open(path, PdfDocumentOpenMode.Import);
            for (int i = 0; i < s.PageCount; i++) outDoc.AddPage(s.Pages[i]);
        }
        using var ms = new MemoryStream();
        outDoc.Save(ms, false);
        return (ms.ToArray(), null);
    }

    /// <summary>Read a consolidated file. scope = Internal | External | Combined; type = excel | pdf. Bumps revision.</summary>
    public static (byte[]? bytes, string fileName, string? err) Download(string platform, string type, string scope)
    {
        lock (Gate)
        {
            bool excel = type.Equals("excel", StringComparison.OrdinalIgnoreCase);
            string ext = excel ? "xlsx" : "pdf";
            string kind = excel ? "MTO" : "Unique";

            if (scope.Equals("Combined", StringComparison.OrdinalIgnoreCase))
            {
                var (bytes, err) = excel ? BuildCombinedExcel(platform) : BuildCombinedPdf(platform);
                if (bytes == null) return (null, "", err);
                int rev = BumpCombinedRev(platform, excel);
                return (bytes, $"{Safe(platform)}_Combined_{kind}_Rev{rev}_{DateTime.Now:yyyy-MM-dd}.{ext}", null);
            }

            string category = scope.Equals("External", StringComparison.OrdinalIgnoreCase) ? "External" : "Internal";
            var path = excel ? XlsxPath(platform, category) : PdfPath(platform, category);
            if (!Have(path))
                return (null, "", excel ? $"No {category} Excel has been added yet." : $"No {category} PDF has been added yet.");

            var m = LoadManifest(platform, category);
            int r = excel ? ++m.ExcelRev : ++m.PdfRev;
            SaveManifest(platform, category, m);
            var b = File.ReadAllBytes(path);
            return (b, $"{Safe(platform)}_{category}_{kind}_Rev{r}_{DateTime.Now:yyyy-MM-dd}.{ext}", null);
        }
    }
}
