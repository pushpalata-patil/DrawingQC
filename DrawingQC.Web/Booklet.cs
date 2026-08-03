using System.Runtime.Versioning;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DrawingQC.Web;

/// <summary>Inputs for one booklet generation run.</summary>
public sealed class BookletInputs
{
    public string TemplatePath { get; set; } = "";   // .docx template
    public string ExcelPath { get; set; } = "";       // .xlsx data
    public string DrawingsPath { get; set; } = "";    // combined drawings .pdf (Appendix B)
    public string? Sheet { get; set; }                // Excel sheet name (default: the BATCH-2 sheet)
    public string? OutputPath { get; set; }           // final merged PDF path (optional)

    // Revision block (fills the cover approval table + revision history + page headers).
    public string Rev { get; set; } = "";
    public string Date { get; set; } = "";
    public string Description { get; set; } = "";     // Description of Revision
    public string Prepared { get; set; } = "";
    public string Verified { get; set; } = "";
    public string Approved { get; set; } = "";        // Approved (Discipline Leader)
}

/// <summary>
/// Assembles the QATAR cable-tray-support booklet:
///  1. Fill Rev/Date and the Appendix A tables (from Excel) into the Word template.
///  2. Convert the filled Word doc to PDF using the locally installed Microsoft Word.
///  3. Append the drawings PDF (Appendix B) after it -> one final booklet PDF.
/// Windows-only: steps 1 &amp; 3 are cross-platform, but step 2 needs Word (COM).
/// </summary>
public static class BookletBuilder
{
    public static string Build(BookletInputs input)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException(
                "Booklet generation runs on the local Windows app (it needs Microsoft Word installed).");

        foreach (var (label, path) in new[] { ("Template", input.TemplatePath), ("Excel", input.ExcelPath), ("Drawings PDF", input.DrawingsPath) })
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                throw new FileNotFoundException($"{label} file not found: {path}");

        string outPdf = string.IsNullOrWhiteSpace(input.OutputPath)
            ? Path.Combine(Path.GetDirectoryName(input.TemplatePath)!,
                Path.GetFileNameWithoutExtension(input.TemplatePath) + " - Booklet.pdf")
            : input.OutputPath!;

        string workDocx = Path.Combine(Path.GetTempPath(), $"booklet_{Guid.NewGuid():N}.docx");
        string bookletPdf = Path.Combine(Path.GetTempPath(), $"booklet_{Guid.NewGuid():N}.pdf");

        try
        {
            File.Copy(input.TemplatePath, workDocx, overwrite: true);

            int filled = FillDocument(workDocx, input);
            if (filled == 0)
                throw new InvalidOperationException("No data rows were read from the Excel sheet.");

            int drawingPages = GetPdfPageCount(input.DrawingsPath);
            ExportPdfWithMeta(workDocx, input.Rev, input.Date, bookletPdf, drawingPages);
            MergePdfs(bookletPdf, input.DrawingsPath, outPdf);
            return outPdf;
        }
        finally
        {
            TryDelete(workDocx);
            TryDelete(bookletPdf);
        }
    }

    // ---------- Step 1: fill Appendix A tables from the Excel ----------

    private static int FillDocument(string docxPath, BookletInputs input)
    {
        var data = ReadExcel(input.ExcelPath, input.Sheet);

        using var doc = WordprocessingDocument.Open(docxPath, isEditable: true);
        var body = doc.MainDocumentPart!.Document.Body!;

        FillCoverRevisionBlock(body, input);

        int idx = 0;
        int appendixTable = 0;
        foreach (var table in body.Descendants<Table>())
        {
            var rows = table.Elements<TableRow>().ToList();
            if (rows.Count == 0) continue;

            // Appendix A page-tables start with a merged "CABLE TRAY SUPPORTS" title row.
            if (!rows[0].InnerText.Contains("CABLE TRAY SUPPORTS", StringComparison.OrdinalIgnoreCase))
                continue;

            appendixTable++;
            SetFaintBorders(table);

            // Force every Appendix A table onto its own page so each page holds exactly
            // 45 entries (the smaller one-line font would otherwise flow two tables per page).
            if (appendixTable > 1)
            {
                var titlePara = rows[0].Elements<TableCell>().FirstOrDefault()?
                    .Elements<Paragraph>().FirstOrDefault();
                if (titlePara != null)
                {
                    titlePara.ParagraphProperties ??= new ParagraphProperties();
                    titlePara.ParagraphProperties.PageBreakBefore = new PageBreakBefore();
                }
            }

            // Rows 0 = title, 1 = column header, 2..n = empty data rows.
            for (int r = 2; r < rows.Count && idx < data.Count; r++)
            {
                var cells = rows[r].Elements<TableCell>().ToList();
                if (cells.Count < 7) continue;

                var entry = data[idx++];
                for (int c = 0; c < 7; c++)
                    SetCellText(cells[c], c < entry.Length ? entry[c] : "", fitOneLine: true);
            }
            if (idx >= data.Count) break;
        }

        FixAppendixBHeading(body);

        doc.MainDocumentPart.Document.Save();
        return idx;
    }

    // The Appendix B heading uses a style ("APPTIT1") the Contents TOC ignores, and its
    // own numbering — so it's missing from the Contents. Copy Appendix A's heading
    // paragraph properties (style "appendix" + its list numbering) onto the Appendix B
    // heading so it becomes the next item ("APPENDIX B") and shows up under Appendix A.
    private static void FixAppendixBHeading(Body body)
    {
        bool IsHeading(Paragraph p, string style, string mustContain) =>
            p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == style &&
            p.InnerText.Contains(mustContain, StringComparison.OrdinalIgnoreCase) &&
            !p.InnerText.Contains("PAGEREF");

        var appA = body.Descendants<Paragraph>()
            .FirstOrDefault(p => IsHeading(p, "appendix", "CABLE TRAY SUPPORT LIST"));
        var appB = body.Descendants<Paragraph>()
            .FirstOrDefault(p => IsHeading(p, "APPTIT1", "DETAILED DRAWING"));

        if (appA?.ParagraphProperties != null && appB != null)
        {
            appB.ParagraphProperties = (ParagraphProperties)appA.ParagraphProperties.CloneNode(true);
            // Start Appendix B at the top of a fresh page so its style's large "space before"
            // is absorbed (otherwise it lands in the middle of the page).
            appB.ParagraphProperties.PageBreakBefore = new PageBreakBefore();
            appB.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "0" };
        }
    }

    private static List<string[]> ReadExcel(string excelPath, string? sheetName)
    {
        using var wb = new XLWorkbook(excelPath);

        IXLWorksheet? ws = null;
        if (!string.IsNullOrWhiteSpace(sheetName))
            ws = wb.Worksheets.FirstOrDefault(s => s.Name.Trim().Equals(sheetName.Trim(), StringComparison.OrdinalIgnoreCase));
        ws ??= wb.Worksheets.FirstOrDefault(s => s.Name.Replace(" ", "").Contains("BATCH-2", StringComparison.OrdinalIgnoreCase));
        ws ??= wb.Worksheets.First();

        var used = ws.RangeUsed();
        var list = new List<string[]>();
        if (used == null) return list;

        int lastRow = used.LastRow().RowNumber();
        // Row 1 = "CABLE TRAY SUPPORTS" title, row 2 = header, data from row 3.
        for (int r = 3; r <= lastRow; r++)
        {
            string slno = ws.Cell(r, 1).GetString().Trim();
            string support = ws.Cell(r, 2).GetString().Trim();
            if (slno.Length == 0 && support.Length == 0) continue;

            var arr = new string[7];
            for (int c = 0; c < 7; c++) arr[c] = ws.Cell(r, c + 1).GetString().Trim();
            list.Add(arr);
        }
        return list;
    }

    // Data font size (half-points): 16 = 8pt, small enough that entries fit on one line.
    private const string DataFontHalfPt = "16";

    // Give an Appendix A table thin, faint, uniform grid lines (and drop any per-cell
    // border overrides so every row/column line looks the same).
    private static void SetFaintBorders(Table table)
    {
        const string faint = "BFBFBF"; // light gray
        var tblPr = table.GetFirstChild<TableProperties>() ?? table.PrependChild(new TableProperties());
        tblPr.RemoveAllChildren<TableBorders>();
        tblPr.Append(new TableBorders(
            new TopBorder { Val = BorderValues.Single, Size = 4, Color = faint },
            new BottomBorder { Val = BorderValues.Single, Size = 4, Color = faint },
            new LeftBorder { Val = BorderValues.Single, Size = 4, Color = faint },
            new RightBorder { Val = BorderValues.Single, Size = 4, Color = faint },
            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = faint },
            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = faint }));

        foreach (var cell in table.Descendants<TableCell>())
            cell.GetFirstChild<TableCellProperties>()?.RemoveAllChildren<TableCellBorders>();
    }

    private static void SetCellText(TableCell cell, string text, bool fitOneLine = false)
    {
        var para = cell.Elements<Paragraph>().FirstOrDefault();
        if (para == null)
        {
            para = new Paragraph();
            cell.Append(para);
        }

        // Preserve any existing run formatting in the cell.
        var runProps = para.Elements<Run>()
            .Select(run => run.RunProperties)
            .FirstOrDefault(rp => rp != null);

        para.RemoveAllChildren<Run>();

        var r = new Run();
        var rPr = runProps != null ? (RunProperties)runProps.CloneNode(true) : new RunProperties();
        if (fitOneLine)
        {
            // Shrink so Appendix A values stay on a single line, and force uniform, regular
            // (non-bold) data — the template left some rows bold and some not.
            rPr.FontSize = new FontSize { Val = DataFontHalfPt };
            rPr.FontSizeComplexScript = new FontSizeComplexScript { Val = DataFontHalfPt };
            rPr.Bold = new Bold { Val = false };
            rPr.BoldComplexScript = new BoldComplexScript { Val = false };
        }
        r.RunProperties = rPr;
        r.Append(new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
        para.Append(r);
    }

    // Fill the cover approval block (Rev/Date/Description/Prepared/Verified/Approved) and
    // the revision-history row (Rev + Revision Description) from the form values.
    private static void FillCoverRevisionBlock(Body body, BookletInputs i)
    {
        foreach (var table in body.Descendants<Table>())
        {
            var rows = table.Elements<TableRow>().ToList();

            // --- Cover approval block: a label row "Rev. | Date | Description of Revision |
            //     Prepared | Verified | Approved(Discipline Leader)" with the data row above it.
            int labelRow = rows.FindIndex(r =>
            {
                var c = r.Elements<TableCell>().Select(x => x.InnerText.Trim()).ToList();
                return c.Count == 6 &&
                       c[0].StartsWith("Rev", StringComparison.OrdinalIgnoreCase) &&
                       c.Any(t => t.Contains("Prepared", StringComparison.OrdinalIgnoreCase));
            });
            if (labelRow > 0)
            {
                var d = rows[labelRow - 1].Elements<TableCell>().ToList();
                if (d.Count == 6)
                {
                    var vals = new[] { i.Rev, i.Date, i.Description, i.Prepared, i.Verified, i.Approved };
                    for (int c = 0; c < 6; c++)
                        if (!string.IsNullOrWhiteSpace(vals[c])) SetCellText(d[c], vals[c]);
                }
            }

            // --- Revision history: header "REVISION | REVISED CHAPTERS | REVISION DESCRIPTION
            //     | REASON FOR REVISION"; fill Rev + Revision Description in the row below.
            int histHdr = rows.FindIndex(r =>
            {
                var c = r.Elements<TableCell>().Select(x => x.InnerText.Trim()).ToList();
                return c.Count == 4 &&
                       c[0].Equals("REVISION", StringComparison.OrdinalIgnoreCase) &&
                       c[2].Contains("DESCRIPTION", StringComparison.OrdinalIgnoreCase);
            });
            if (histHdr >= 0 && histHdr + 1 < rows.Count)
            {
                var h = rows[histHdr + 1].Elements<TableCell>().ToList();
                if (h.Count == 4)
                {
                    if (!string.IsNullOrWhiteSpace(i.Rev)) SetCellText(h[0], i.Rev);
                    if (!string.IsNullOrWhiteSpace(i.Description)) SetCellText(h[2], i.Description);
                }
            }
        }
    }

    // ---------- Step 2: Rev/Date replace + DOCX -> PDF via Word COM ----------

    private static int GetPdfPageCount(string path)
    {
        try
        {
            using var d = PdfReader.Open(path, PdfDocumentOpenMode.InformationOnly);
            return d.PageCount;
        }
        catch { return 0; }
    }

    [SupportedOSPlatform("windows")]
    private static void ExportPdfWithMeta(string docxPath, string rev, string date, string pdfPath, int extraPages)
    {
        Type? wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
            throw new InvalidOperationException("Microsoft Word is not installed on this machine.");

        dynamic app = Activator.CreateInstance(wordType)!;
        dynamic? document = null;
        try
        {
            app.Visible = false;
            app.DisplayAlerts = 0; // wdAlertsNone

            document = app.Documents.Open(docxPath, ReadOnly: false, Visible: false);

            // Fill Rev / Date wherever the template shows them (body + all section headers/footers).
            if (!string.IsNullOrWhiteSpace(date))
            {
                ReplaceEverywhere(document, "24-07-2026", date);
                ReplaceEverywhere(document, "24/07/2026", date);
            }
            if (!string.IsNullOrWhiteSpace(rev))
            {
                ReplaceEverywhere(document, "Rev. 00", "Rev. " + rev);
                ReplaceEverywhere(document, "REV 00", "REV " + rev);
                ReplaceEverywhere(document, "rev-00", "rev-" + rev);
            }

            // Replace the template's hardcoded sheet total ("of 19") with the real total:
            // booklet pages + appended drawing pages.
            try
            {
                // Count only the booklet pages (cover..Appendix B), NOT the appended
                // support drawings — the sheet total stops when the index/list is over.
                int total = (int)document.ComputeStatistics(2); // wdStatisticPages
                ReplaceEverywhere(document, "of 19", "of " + total);
            }
            catch { }

            // Refresh the Table(s) of Contents so page numbers are correct and the
            // Appendix A/B entries are all listed, then update remaining fields.
            try { foreach (dynamic toc in document.TablesOfContents) toc.Update(); } catch { }
            try { document.Fields.Update(); } catch { }

            document.ExportAsFixedFormat(pdfPath, 17); // wdExportFormatPDF
        }
        finally
        {
            try { document?.Close(false); } catch { }
            try { app.Quit(); } catch { }
        }
    }

    [SupportedOSPlatform("windows")]
    private static void ReplaceEverywhere(dynamic document, string find, string replace)
    {
        ReplaceInRange(document.Content, find, replace);
        foreach (dynamic section in document.Sections)
        {
            foreach (dynamic h in section.Headers) ReplaceInRange(h.Range, find, replace);
            foreach (dynamic f in section.Footers) ReplaceInRange(f.Range, find, replace);
        }
    }

    [SupportedOSPlatform("windows")]
    private static void ReplaceInRange(dynamic range, string find, string replace)
    {
        try
        {
            dynamic f = range.Find;
            f.ClearFormatting();
            f.Replacement.ClearFormatting();
            // Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
            //   MatchAllWordForms, Forward, Wrap(1=wdFindContinue), Format, ReplaceWith, Replace(2=wdReplaceAll))
            f.Execute(find, false, false, false, false, false, true, 1, false, replace, 2);
        }
        catch { /* a story without a Find range is fine to skip */ }
    }

    // ---------- Step 3: merge booklet PDF + drawings PDF ----------

    private static void MergePdfs(string bookletPdf, string drawingsPdf, string finalPdf)
    {
        using var outDoc = new PdfDocument();
        AppendPages(outDoc, bookletPdf);
        AppendPages(outDoc, drawingsPdf);
        outDoc.Save(finalPdf);
    }

    private static void AppendPages(PdfDocument outDoc, string pdfPath)
    {
        using var src = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import);
        for (int i = 0; i < src.PageCount; i++)
            outDoc.AddPage(src.Pages[i]);
    }

    private static void TryDelete(string path)
    {
        try { if (File.Exists(path)) File.Delete(path); } catch { }
    }
}
