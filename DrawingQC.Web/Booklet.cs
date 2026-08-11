using System.Runtime.Versioning;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace DrawingQC.Web;

/// <summary>Result of a booklet run: the merged PDF and the editable Word booklet.</summary>
public sealed record BookletResult(string PdfPath, string DocxPath);

/// <summary>Inputs for one booklet generation run.</summary>
public sealed class BookletInputs
{
    public string TemplatePath { get; set; } = "";   // .docx template
    public string ExcelPath { get; set; } = "";       // .xlsx support-list data
    public string DrawingsPath { get; set; } = "";    // combined drawings .pdf (Appendix B)
    public string? BomPath { get; set; }              // optional BOM .xlsx (Material Statistics table)
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
    public static BookletResult Build(BookletInputs input)
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
        // The Word download is the booklet body (cover..Appendix B) as an editable .docx,
        // written alongside the PDF. The appended drawings stay PDF-only (they can't merge
        // into Word), so the .docx omits them.
        string outDocx = Path.ChangeExtension(outPdf, ".docx");

        string workDocx = Path.Combine(Path.GetTempPath(), $"booklet_{Guid.NewGuid():N}.docx");
        string bookletPdf = Path.Combine(Path.GetTempPath(), $"booklet_{Guid.NewGuid():N}.pdf");

        try
        {
            File.Copy(input.TemplatePath, workDocx, overwrite: true);

            int filled = FillDocument(workDocx, input);
            if (filled == 0)
                throw new InvalidOperationException("No data rows were read from the Excel sheet.");

            int drawingPages = GetPdfPageCount(input.DrawingsPath);
            // ExportPdfWithMeta also saves the filled+meta docx back to workDocx, so we can
            // hand it out as the Word download.
            ExportPdfWithMeta(workDocx, input.Rev, input.Date, bookletPdf, drawingPages);
            File.Copy(workDocx, outDocx, overwrite: true);
            MergePdfs(bookletPdf, input.DrawingsPath, outPdf);
            return new BookletResult(outPdf, outDocx);
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

        // Book Statistics: "Total number of Supports" = the row count from the support list.
        FillTotalSupports(body, data.Count);

        // Optional: fill the "Material Statistics for Released Supports" table from a BOM Excel.
        if (!string.IsNullOrWhiteSpace(input.BomPath) && File.Exists(input.BomPath))
            FillMaterialStatistics(body, input.BomPath!);

        // Fill the Appendix A support list. Older templates ship pre-made "CABLE TRAY
        // SUPPORTS" tables; newer templates have none, so we generate them from the data.
        var listTables = body.Descendants<Table>()
            .Where(t => t.Elements<TableRow>().FirstOrDefault()?
                .InnerText.Contains("CABLE TRAY SUPPORTS", StringComparison.OrdinalIgnoreCase) == true)
            .ToList();

        // The index (list) pages get NO repeating title-block header, so size 45 rows to the
        // full body height (no header subtraction). Column widths fit the printable width.
        int pageBudget = ComputePageBudget(doc, body, includeHeader: false);
        int[] colWidths = ScaleColumns(ComputeUsableWidth(body));

        int filled;
        if (listTables.Count > 0)
        {
            // Shrink the empty paragraphs between tables to ~1pt so tall page-filling rows
            // don't spill onto (and create) blank pages.
            ShrinkInterTableBlanks(listTables);
            filled = FillPreMadeTables(listTables, data, pageBudget, colWidths);
        }
        else
        {
            filled = GenerateListTables(body, data, pageBudget, colWidths);
        }

        // Put the list in its own section with a blank header so the title-block header does not
        // appear on the index pages (kept on the cover/section/appendix pages).
        try { RemoveListHeader(doc, body); } catch { /* keep the header if section surgery fails */ }

        // Newer 3-appendix template: drop the "Standard Details" appendix so the
        // "Detailed Drawings" appendix becomes B (matching the old 2-appendix format).
        RemoveStandardDetailsSection(body);

        // Make the "Detailed Drawings" heading appear as the last appendix (B), top of page.
        FixDrawingsAppendixHeading(body);

        doc.MainDocumentPart.Document.Save();
        return filled;
    }

    // Shrink the empty paragraphs the template inserts between the list tables to ~1pt.
    // (Removing them would let adjacent tables merge into one giant table, which paginates
    // chaotically; keeping tiny separators keeps tables distinct without wasting space.)
    private static void ShrinkInterTableBlanks(List<Table> tables)
    {
        if (tables.Count < 2) return;
        var last = tables[tables.Count - 1];
        for (var e = tables[0].NextSibling(); e != null && e != last; e = e.NextSibling())
        {
            if (e is not Paragraph p || p.InnerText.Trim().Length != 0) continue;
            if (p.ParagraphProperties?.SectionProperties != null) continue;

            var pPr = p.ParagraphProperties ??= new ParagraphProperties();
            pPr.SpacingBetweenLines = new SpacingBetweenLines
            { Before = "0", After = "0", Line = "20", LineRule = LineSpacingRuleValues.Exact };
            var mark = pPr.ParagraphMarkRunProperties ??= new ParagraphMarkRunProperties();
            mark.RemoveAllChildren<FontSize>();
            mark.AppendChild(new FontSize { Val = "2" }); // 1pt
        }
    }

    // Resize a pre-made list table's columns to fit the template's printable width.
    private static void SetListColumnWidths(Table table, int[] colWidths)
    {
        var grid = table.GetFirstChild<TableGrid>();
        if (grid != null)
        {
            var cols = grid.Elements<GridColumn>().ToList();
            for (int c = 0; c < cols.Count && c < colWidths.Length; c++)
                cols[c].Width = colWidths[c].ToString();
        }

        var tblPr = table.GetFirstChild<TableProperties>() ?? table.PrependChild(new TableProperties());
        tblPr.RemoveAllChildren<TableLayout>();
        tblPr.Append(new TableLayout { Type = TableLayoutValues.Fixed });

        foreach (var row in table.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            if (cells.Count != colWidths.Length) continue; // skip the merged title row
            for (int c = 0; c < cells.Count; c++)
            {
                var tcPr = cells[c].GetFirstChild<TableCellProperties>()
                    ?? cells[c].PrependChild(new TableCellProperties());
                tcPr.RemoveAllChildren<TableCellWidth>();
                tcPr.PrependChild(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = colWidths[c].ToString() });
            }
        }
    }

    // Twips available for data rows on one list page = usable body height (page height minus
    // margins minus the tall repeating page header) minus the table's title + header rows.
    // Deriving this from the actual template geometry keeps 45 rows filling exactly one page
    // regardless of page size / header height (a fixed 315 overflowed header-heavy templates).
    private static int ComputePageBudget(WordprocessingDocument doc, Body body, bool includeHeader = true)
    {
        int pageH = 16838, topM = 1440, bottomM = 1440, headerDist = 720; // A4 defaults
        var sect = body.Descendants<SectionProperties>().FirstOrDefault();
        if (sect != null)
        {
            var ps = sect.GetFirstChild<PageSize>();
            var pm = sect.GetFirstChild<PageMargin>();
            if (ps?.Height?.Value != null) pageH = (int)ps.Height.Value;
            if (pm != null)
            {
                if (pm.Top?.Value != null) topM = pm.Top.Value;
                if (pm.Bottom?.Value != null) bottomM = pm.Bottom.Value;
                if (pm.Header?.Value != null) headerDist = (int)pm.Header.Value;
            }
        }
        // Tallest repeating page header — its explicit table-row heights push the body down.
        int headerH = 0;
        if (includeHeader)
            foreach (var hp in doc.MainDocumentPart!.HeaderParts)
            {
                int sum = 0;
                foreach (var tr in hp.Header.Descendants<TableRow>())
                {
                    var trh = tr.TableRowProperties?.GetFirstChild<TableRowHeight>();
                    sum += trh?.Val?.Value != null ? (int)trh.Val.Value : 300;
                }
                headerH = Math.Max(headerH, sum);
            }
        // Header content (logos) usually renders taller than its explicit row heights, so add a
        // buffer; otherwise 45 rows overflow by a few and spill onto a short continuation page.
        // When the list section has no header (includeHeader=false) the body starts at the margin.
        int bodyTop = includeHeader ? Math.Max(topM, headerDist + (int)(headerH * 1.35)) : topM;
        int forData = pageH - bodyTop - bottomM - 750 /* table title + header rows */ - 300 /* safety */;
        return Math.Clamp(forData, 6000, 15000);
    }

    // Printable width (page width minus left/right margins) for THIS template.
    private static int ComputeUsableWidth(Body body)
    {
        int pageW = 11906, leftM = 1440, rightM = 1440; // A4 defaults
        var sect = body.Descendants<SectionProperties>().FirstOrDefault();
        if (sect != null)
        {
            var ps = sect.GetFirstChild<PageSize>();
            var pm = sect.GetFirstChild<PageMargin>();
            if (ps?.Width?.Value != null) pageW = (int)ps.Width.Value;
            if (pm != null)
            {
                if (pm.Left?.Value != null) leftM = (int)pm.Left.Value;
                if (pm.Right?.Value != null) rightM = (int)pm.Right.Value;
            }
        }
        return Math.Max(4000, pageW - leftM - rightM);
    }

    // Minimum width each column needs so its 9pt text stays on one line (S.L., SUPPORT No.,
    // DRAWING No., REVISION, LEVEL, PRESENT STATUS, REMARKS). Used when the printable width is
    // narrower than the reference table so the data columns don't shrink enough to wrap/clip.
    private static readonly int[] ListColMins = { 480, 1680, 1680, 1000, 1600, 1550, 850 };

    // Fit the reference column widths into the template's printable width. A wider page scales
    // up proportionally; a narrower page shrinks each column toward its minimum in proportion
    // to its slack, so the (empty) Present Status / Remarks columns give up space first and the
    // data columns stay wide enough for their text.
    private static int[] ScaleColumns(int usableWidth)
    {
        int refSum = ListColWidths.Sum();
        var scaled = new int[ListColWidths.Length];

        if (usableWidth >= refSum)
        {
            int acc = 0;
            for (int i = 0; i < scaled.Length; i++) { scaled[i] = (int)Math.Round((double)ListColWidths[i] * usableWidth / refSum); acc += scaled[i]; }
            scaled[^1] += usableWidth - acc;
            return scaled;
        }

        int minSum = ListColMins.Sum();
        if (usableWidth <= minSum) // extremely narrow: scale the minimums (last resort)
        {
            int acc = 0;
            for (int i = 0; i < scaled.Length; i++) { scaled[i] = (int)Math.Round((double)ListColMins[i] * usableWidth / minSum); acc += scaled[i]; }
            scaled[^1] += usableWidth - acc;
            return scaled;
        }

        int deficit = refSum - usableWidth;
        int totalSlack = refSum - minSum;
        int used = 0;
        for (int i = 0; i < scaled.Length; i++)
        {
            int reduce = (int)Math.Round((double)deficit * (ListColWidths[i] - ListColMins[i]) / totalSlack);
            scaled[i] = ListColWidths[i] - reduce;
            used += scaled[i];
        }
        scaled[^1] += usableWidth - used; // absorb rounding drift
        return scaled;
    }

    // Order a sectPr's SectionType correctly (it must sit just before PageSize).
    private static void SetNextPageType(SectionProperties sect)
    {
        sect.RemoveAllChildren<SectionType>();
        var type = new SectionType { Val = SectionMarkValues.NextPage };
        var pgSz = sect.GetFirstChild<PageSize>();
        if (pgSz != null) sect.InsertBefore(type, pgSz); else sect.AppendChild(type);
    }

    // Put the support list in its own section that references an EMPTY header, so the repeating
    // title-block header does not print on the index pages. The pages before (cover/sections/
    // appendix divider) and after (detailed drawings) keep their normal header.
    private static void RemoveListHeader(WordprocessingDocument doc, Body body)
    {
        var main = doc.MainDocumentPart!;
        var mainSect = body.Elements<SectionProperties>().LastOrDefault();
        if (mainSect == null) return;

        static bool IsListTable(Table t) => t.Elements<TableRow>().FirstOrDefault()?
            .InnerText.Contains("CABLE TRAY SUPPORTS", StringComparison.OrdinalIgnoreCase) == true;

        var firstListTable = body.Descendants<Table>().FirstOrDefault(IsListTable);
        var lastListTable = body.Descendants<Table>().LastOrDefault(IsListTable);
        if (firstListTable == null || lastListTable == null) return;

        // The paragraph just before the first list table ends the header'd section (section A).
        Paragraph? before = null;
        for (var e = firstListTable.PreviousSibling(); e != null; e = e.PreviousSibling())
            if (e is Paragraph p) { before = p; break; }
        if (before == null) return;

        var sectA = (SectionProperties)mainSect.CloneNode(true); // keeps the title-block header
        SetNextPageType(sectA);
        var bpr = before.ParagraphProperties ??= new ParagraphProperties();
        bpr.RemoveAllChildren<SectionProperties>();
        bpr.Append(sectA);

        // Empty header referenced by the list section.
        var emptyHeader = main.AddNewPart<HeaderPart>();
        emptyHeader.Header = new Header();
        emptyHeader.Header.Save();
        string emptyId = main.GetIdOfPart(emptyHeader);

        var sectB = (SectionProperties)mainSect.CloneNode(true);
        sectB.RemoveAllChildren<HeaderReference>();
        SetNextPageType(sectB);
        foreach (var t in new[] { HeaderFooterValues.Default, HeaderFooterValues.Even, HeaderFooterValues.First })
            sectB.InsertAt(new HeaderReference { Type = t, Id = emptyId }, 0);

        // A trailing paragraph after the last list table carries the blank-header list section.
        lastListTable.InsertAfterSelf(new Paragraph(new ParagraphProperties(sectB)));
    }

    // Old template: fill the pre-made list tables in document order. Every physical data row
    // is packed with the next sequential entry (no blank rows). When the data runs out, the
    // leftover empty rows are deleted and any wholly-unused trailing tables are removed, so
    // there are no blank gaps and the total page count drops accordingly.
    private static int FillPreMadeTables(List<Table> tables, List<string[]> data, int pageBudget, int[] colWidths)
    {
        int idx = 0, tableNo = 0;
        foreach (var table in tables)
        {
            var rows = table.Elements<TableRow>().ToList();

            // Data already exhausted on an earlier table: this whole table is unused -> drop it.
            if (idx >= data.Count)
            {
                table.Remove();
                continue;
            }

            tableNo++;
            SetFaintBorders(table);
            SetListColumnWidths(table, colWidths);

            // Each table starts on its own page (the blank inter-table paragraphs that used
            // to cause spill/blank pages were removed).
            if (tableNo > 1)
            {
                var titlePara = rows[0].Elements<TableCell>().FirstOrDefault()?
                    .Elements<Paragraph>().FirstOrDefault();
                if (titlePara != null)
                {
                    titlePara.ParagraphProperties ??= new ParagraphProperties();
                    titlePara.ParagraphProperties.PageBreakBefore = new PageBreakBefore();
                }
            }

            // Row height adapts to the table's real row count so all its rows fill one page.
            int dataRowCount = Math.Max(1, rows.Count - 2);
            int rowH = Math.Clamp(pageBudget / dataRowCount, 200, 340);

            int r = 2;
            for (; r < rows.Count && idx < data.Count; r++)
            {
                var cells = rows[r].Elements<TableCell>().ToList();
                if (cells.Count < 7) continue;
                rows[r].TableRowProperties = new TableRowProperties(
                    new TableRowHeight { Val = (uint)rowH, HeightType = HeightRuleValues.Exact });
                var entry = data[idx++];
                for (int c = 0; c < 7; c++)
                    SetCellText(cells[c], c < entry.Length ? entry[c] : "", fitOneLine: true);
            }

            // Data exhausted partway through this table: delete the leftover empty rows so the
            // final page ends cleanly instead of showing a block of blank rows.
            if (idx >= data.Count)
                for (int k = rows.Count - 1; k >= r; k--)
                    rows[k].Remove();
        }

        // The template's pre-made tables have a fixed capacity (this one holds 563 supports).
        // If the Excel has MORE entries than that, the extras would otherwise be silently
        // dropped.
        if (idx < data.Count && tables.Count > 0)
        {
            // First, keep filling the LAST pre-made table (which usually has spare room on its
            // page) by cloning its rows, up to a full page — so the extras continue on the same
            // page instead of opening a new one.
            ExtendTable(tables[tables.Count - 1], data, ref idx, RowsPerPage, pageBudget);

            // Only if entries still remain do we add further full pages, cloned from a pre-made
            // table so the header/style is identical.
            var template = tables[0];
            OpenXmlElement anchor = tables[tables.Count - 1];
            while (idx < data.Count)
            {
                var clone = (Table)template.CloneNode(true);
                var crows = clone.Elements<TableRow>().ToList();

                var titlePara = crows[0].Elements<TableCell>().FirstOrDefault()?
                    .Elements<Paragraph>().FirstOrDefault();
                if (titlePara != null)
                {
                    titlePara.ParagraphProperties ??= new ParagraphProperties();
                    titlePara.ParagraphProperties.PageBreakBefore = new PageBreakBefore();
                }

                int dataRowCount = Math.Max(1, crows.Count - 2);
                int rowH = Math.Clamp(pageBudget / dataRowCount, 200, 340);

                int r = 2;
                for (; r < crows.Count && idx < data.Count; r++)
                {
                    var cells = crows[r].Elements<TableCell>().ToList();
                    if (cells.Count < 7) continue;
                    crows[r].TableRowProperties = new TableRowProperties(
                        new TableRowHeight { Val = (uint)rowH, HeightType = HeightRuleValues.Exact });
                    var entry = data[idx++];
                    for (int c = 0; c < 7; c++)
                        SetCellText(cells[c], c < entry.Length ? entry[c] : "", fitOneLine: true);
                }
                if (idx >= data.Count)
                    for (int k = crows.Count - 1; k >= r; k--)
                        crows[k].Remove();

                anchor.InsertAfterSelf(clone);
                anchor = clone;
            }
        }

        return idx;
    }

    // Append more data rows to an existing list table (cloning its last data row for identical
    // structure) until the data runs out or the table reaches maxDataRows. Row heights are then
    // recomputed uniformly so all rows still fit on one page.
    private static void ExtendTable(Table table, List<string[]> data, ref int idx, int maxDataRows, int pageBudget)
    {
        var dataRows = table.Elements<TableRow>().Skip(2).ToList(); // skip title + header
        if (dataRows.Count == 0) return;
        var sample = dataRows[dataRows.Count - 1];
        int current = dataRows.Count;

        while (idx < data.Count && current < maxDataRows)
        {
            var newRow = (TableRow)sample.CloneNode(true);
            var cells = newRow.Elements<TableCell>().ToList();
            if (cells.Count < 7) break;
            var entry = data[idx++];
            for (int c = 0; c < 7; c++)
                SetCellText(cells[c], c < entry.Length ? entry[c] : "", fitOneLine: true);
            table.Append(newRow);
            current++;
        }

        int rowH = Math.Clamp(pageBudget / Math.Max(1, current), 200, 340);
        foreach (var row in table.Elements<TableRow>().Skip(2))
            row.TableRowProperties = new TableRowProperties(
                new TableRowHeight { Val = (uint)rowH, HeightType = HeightRuleValues.Exact });
    }

    private static readonly string[] ListHeaders =
        { "S.L.", "SUPPORT No.", "DRAWING No.", "REVISION", "LEVEL", "PRESENT STATUS", "REMARKS" };
    // Column widths tuned so (a) 14-char HK-CTS-…-#### tags fit on one line at 9pt in the
    // SUPPORT No. / DRAWING No. columns, and (b) the "PRESENT STATUS" header stays on one
    // line. Extra width for those comes from the mostly-empty REMARKS column. Total unchanged.
    private static readonly int[] ListColWidths = { 557, 1720, 1720, 1148, 2015, 2240, 1172 };
    private const int RowsPerPage = 45;        // data entries per page

    // New template: build the Appendix A list tables (45 rows/page) from the Excel and
    // insert them right after the "CABLE TRAY SUPPORT LIST" heading.
    private static int GenerateListTables(Body body, List<string[]> data, int pageBudget, int[] colWidths)
    {
        var heading = body.Descendants<Paragraph>().FirstOrDefault(p =>
            p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "appendix" &&
            p.InnerText.Contains("CABLE TRAY SUPPORT LIST", StringComparison.OrdinalIgnoreCase) &&
            !p.InnerText.Contains("PAGEREF"));
        if (heading == null) return 0;

        // Row height so RowsPerPage rows fill exactly one page of THIS template.
        int rowHeight = Math.Clamp(pageBudget / RowsPerPage, 200, 340);

        OpenXmlElement anchor = heading;
        int idx = 0;
        bool first = true;
        while (idx < data.Count)
        {
            var chunk = data.Skip(idx).Take(RowsPerPage).ToList();
            idx += chunk.Count;
            // The first table needs no page break — the list section break already starts it on a
            // fresh page (below the APPENDIX A divider); the rest start their own page.
            var table = BuildListTable(chunk, pageBreakBefore: !first, rowHeight, colWidths);
            first = false;
            anchor.InsertAfterSelf(table);
            anchor = table;
        }
        return idx;
    }

    private static Table BuildListTable(List<string[]> chunk, bool pageBreakBefore, int rowHeight, int[] colWidths)
    {
        const string faint = "BFBFBF";
        int totalW = colWidths.Sum();

        var table = new Table(
            new TableProperties(
                new TableWidth { Type = TableWidthUnitValues.Dxa, Width = totalW.ToString() },
                new TableLayout { Type = TableLayoutValues.Fixed },
                // Small cell padding so narrow columns (S.L., REVISION, REMARKS) fit their text
                // on one line instead of wrapping and being clipped by the exact row height.
                new TableCellMarginDefault(
                    new TableCellLeftMargin { Width = 45, Type = TableWidthValues.Dxa },
                    new TableCellRightMargin { Width = 45, Type = TableWidthValues.Dxa }),
                new TableBorders(
                    new TopBorder { Val = BorderValues.Single, Size = 4, Color = faint },
                    new BottomBorder { Val = BorderValues.Single, Size = 4, Color = faint },
                    new LeftBorder { Val = BorderValues.Single, Size = 4, Color = faint },
                    new RightBorder { Val = BorderValues.Single, Size = 4, Color = faint },
                    new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4, Color = faint },
                    new InsideVerticalBorder { Val = BorderValues.Single, Size = 4, Color = faint })));

        var grid = new TableGrid();
        foreach (var w in colWidths) grid.Append(new GridColumn { Width = w.ToString() });
        table.Append(grid);

        // Title row: one merged cell spanning all 7 columns.
        var titlePara = new Paragraph(
            new ParagraphProperties(new Justification { Val = JustificationValues.Center }),
            new Run(new RunProperties(new Bold()), new Text("CABLE TRAY SUPPORTS")));
        if (pageBreakBefore) titlePara.ParagraphProperties!.PageBreakBefore = new PageBreakBefore();
        table.Append(new TableRow(new TableCell(
            new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = totalW.ToString() },
                new GridSpan { Val = 7 },
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
            titlePara)));

        // Header row. 8pt bold keeps the long headers ("PRESENT STATUS", "REVISION", "REMARKS")
        // and the narrow "S.L." on one line in these column widths (the default style is larger
        // and wraps). Used only for the new-template path; the old template clones its own rows.
        var headerRow = new TableRow();
        for (int c = 0; c < 7; c++)
            headerRow.Append(MakeCell(ListHeaders[c], colWidths[c], bold: true, center: true, font: "16"));
        table.Append(headerRow);

        // Data rows (9pt, S.L. + REVISION centred), tall enough to fill the page.
        foreach (var entry in chunk)
        {
            var row = new TableRow(new TableRowProperties(
                new TableRowHeight { Val = (uint)rowHeight, HeightType = HeightRuleValues.Exact }));
            for (int c = 0; c < 7; c++)
                row.Append(MakeCell(c < entry.Length ? entry[c] : "", colWidths[c],
                    bold: false, center: c == 0 || c == 3, font: DataFontHalfPt));
            table.Append(row);
        }
        return table;
    }

    private static TableCell MakeCell(string text, int widthTwips, bool bold, bool center, string? font)
    {
        var rPr = new RunProperties();
        if (bold) rPr.Bold = new Bold();
        if (font != null)
        {
            rPr.FontSize = new FontSize { Val = font };
            rPr.FontSizeComplexScript = new FontSizeComplexScript { Val = font };
        }
        var run = new Run(new Text(text ?? "") { Space = SpaceProcessingModeValues.Preserve });
        if (rPr.HasChildren) run.RunProperties = rPr;

        var pPr = new ParagraphProperties(
            // Compact rows (no space before/after, single line) so 45 fit on one page.
            new SpacingBetweenLines { Before = "0", After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto });
        if (center) pPr.Justification = new Justification { Val = JustificationValues.Center };

        return new TableCell(
            new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = widthTwips.ToString() },
                new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }),
            new Paragraph(pPr, run));
    }

    // Remove the whole "Cable Tray Support Standard Details" appendix (heading + content)
    // when present, so "Detailed Drawings" renumbers from C to B.
    private static void RemoveStandardDetailsSection(Body body)
    {
        var kids = body.ChildElements.ToList();
        bool IsAppendix(OpenXmlElement e, string contains) =>
            e is Paragraph p &&
            p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "appendix" &&
            p.InnerText.Contains(contains, StringComparison.OrdinalIgnoreCase) &&
            !p.InnerText.Contains("PAGEREF");

        int start = kids.FindIndex(e => IsAppendix(e, "STANDARD DETAILS"));
        if (start < 0) return;

        int end = kids.FindIndex(start + 1, e => IsAppendix(e, "DETAILED DRAWING"));
        if (end < 0) return;

        for (int i = end - 1; i >= start; i--) kids[i].Remove();
    }

    // Make the "Detailed Drawings" heading the last appendix and start it at the top of a
    // fresh page. Old template: its heading uses style "APPTIT1" (which the Contents TOC
    // ignores), so copy Appendix A's "appendix" style/numbering onto it. New template: it
    // is already "appendix" (renumbers to B once Standard Details is removed).
    private static void FixDrawingsAppendixHeading(Body body)
    {
        bool IsHeading(Paragraph p, string mustContain, params string[] styles) =>
            styles.Contains(p.ParagraphProperties?.ParagraphStyleId?.Val?.Value) &&
            p.InnerText.Contains(mustContain, StringComparison.OrdinalIgnoreCase) &&
            !p.InnerText.Contains("PAGEREF");

        var appA = body.Descendants<Paragraph>()
            .FirstOrDefault(p => IsHeading(p, "CABLE TRAY SUPPORT LIST", "appendix"));
        var draw = body.Descendants<Paragraph>()
            .FirstOrDefault(p => IsHeading(p, "DETAILED DRAWING", "appendix", "APPTIT1"));
        if (draw == null) return;

        var style = draw.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (style == "APPTIT1" && appA?.ParagraphProperties != null)
            draw.ParagraphProperties = (ParagraphProperties)appA.ParagraphProperties.CloneNode(true);

        draw.ParagraphProperties ??= new ParagraphProperties();
        draw.ParagraphProperties.PageBreakBefore = new PageBreakBefore();
        draw.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines { Before = "0", After = "0" };
    }

    // Set "Total number of Supports - N" in the Book Statistics section to the actual count.
    private static void FillTotalSupports(Body body, int count)
    {
        foreach (var p in body.Descendants<Paragraph>())
        {
            var txt = p.InnerText;
            if (!txt.Contains("Total number of Supports", StringComparison.OrdinalIgnoreCase)) continue;
            if (txt.Contains("HOLD", StringComparison.OrdinalIgnoreCase)) continue; // the "under HOLD" line
            if (txt.Contains("PAGEREF")) continue;

            // Replace the number in the last run that has digits.
            var runs = p.Descendants<Text>().ToList();
            for (int i = runs.Count - 1; i >= 0; i--)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(runs[i].Text, @"\d"))
                {
                    runs[i].Text = System.Text.RegularExpressions.Regex.Replace(
                        runs[i].Text, @"\d[\d,]*", count.ToString());
                    runs[i].Space = SpaceProcessingModeValues.Preserve;
                    return;
                }
            }
        }
    }

    // Fill the "Material Statistics for Released Supports" table from a BOM Excel
    // (columns Description | Material | Length/Area | Weight; SL No auto-numbered; the
    // Total row's weight is taken from the BOM).
    private static void FillMaterialStatistics(Body body, string bomPath)
    {
        var (data, totalWeight) = ReadBom(bomPath);
        if (data.Count == 0) return;

        var table = body.Descendants<Table>().FirstOrDefault(t =>
            t.InnerText.Contains("WEIGHT", StringComparison.OrdinalIgnoreCase) &&
            t.InnerText.Contains("MATERIAL", StringComparison.OrdinalIgnoreCase) &&
            t.InnerText.Contains("DESCRIPTION", StringComparison.OrdinalIgnoreCase));
        if (table == null) return;

        var rows = table.Elements<TableRow>().ToList();
        if (rows.Count < 2) return;

        var templateData = (TableRow)rows[1].CloneNode(true);           // row formatting template
        var templateTotal = (TableRow)rows[rows.Count - 1].CloneNode(true); // total-row template

        // Keep only the header row.
        for (int i = rows.Count - 1; i >= 1; i--) rows[i].Remove();

        // Material rows.
        for (int i = 0; i < data.Count; i++)
        {
            var row = (TableRow)templateData.CloneNode(true);
            var cells = row.Elements<TableCell>().ToList();
            if (cells.Count >= 5)
            {
                SetCellText(cells[0], (i + 1).ToString(), center: true);
                SetCellText(cells[1], data[i][0], center: true); // description
                SetCellText(cells[2], data[i][1], center: true); // material
                SetCellText(cells[3], data[i][2], center: true); // length / area
                SetCellText(cells[4], data[i][3], center: true); // weight
            }
            table.Append(row);
        }

        // Total row: a "TOTAL" label next to the summed weight (both centred to line up).
        var trow = (TableRow)templateTotal.CloneNode(true);
        var tcells = trow.Elements<TableCell>().ToList();
        if (tcells.Count >= 5 && !string.IsNullOrWhiteSpace(totalWeight))
        {
            SetCellText(tcells[3], "TOTAL", center: true);
            SetCellText(tcells[4], totalWeight, center: true);
        }
        table.Append(trow);

        KeepNoteWithTable(body, table);
    }

    // Keep the "MTO not inclusive…" note directly under the Material Statistics table (stop it
    // spilling to the next page): drop the blank paragraph(s) between them and mark the table's
    // paragraphs keep-with-next so Word keeps the table + note together on one page.
    private static void KeepNoteWithTable(Body body, Table table)
    {
        var note = body.Descendants<Paragraph>()
            .FirstOrDefault(p => p.InnerText.Contains("MTO not inclusive", StringComparison.OrdinalIgnoreCase));

        for (var e = table.NextSibling(); e != null && e != note;)
        {
            var next = e.NextSibling();
            if (e is Paragraph p && p.InnerText.Trim().Length == 0 && p.ParagraphProperties?.SectionProperties == null)
                e.Remove();
            e = next;
        }

        foreach (var para in table.Descendants<Paragraph>())
        {
            para.ParagraphProperties ??= new ParagraphProperties();
            para.ParagraphProperties.KeepNext ??= new KeepNext();
        }
        if (note != null)
        {
            // Remove blank paragraphs between the note and the next "appendix" heading — otherwise
            // one lands alone on a page and creates a blank divider page before Appendix A.
            for (var e = note.NextSibling(); e != null;)
            {
                var next = e.NextSibling();
                if (e is Paragraph ap && ap.ParagraphProperties?.ParagraphStyleId?.Val?.Value == "appendix") break;
                if (e is Paragraph bp && bp.InnerText.Trim().Length == 0 && bp.ParagraphProperties?.SectionProperties == null)
                    e.Remove();
                e = next;
            }
            note.ParagraphProperties ??= new ParagraphProperties();
            note.ParagraphProperties.KeepLines ??= new KeepLines();
        }
    }

    // Round a numeric string to N decimals (removes float noise like 2063.2100000000028).
    private static string CleanNumber(string s, int decimals)
    {
        if (double.TryParse(s, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double d))
            return d.ToString("F" + decimals, System.Globalization.CultureInfo.InvariantCulture);
        return s;
    }

    // Read a BOM Excel -> material rows [Description, Material, Length/Area, Weight] + total weight.
    private static (List<string[]> rows, string totalWeight) ReadBom(string path)
    {
        var rows = new List<string[]>();
        string totalWeight = "";

        using var wb = new XLWorkbook(path);
        var ws = wb.Worksheets.FirstOrDefault(s => s.RangeUsed() != null) ?? wb.Worksheets.First();
        var used = ws.RangeUsed();
        if (used == null) return (rows, totalWeight);

        int firstRow = used.FirstRow().RowNumber(), lastRow = used.LastRow().RowNumber();
        int firstCol = used.FirstColumn().ColumnNumber(), lastCol = used.LastColumn().ColumnNumber();

        // Locate the header row and column positions by matching header names.
        int headerRow = -1, cDesc = -1, cMat = -1, cLen = -1, cWt = -1;
        for (int r = firstRow; r <= Math.Min(firstRow + 10, lastRow) && headerRow < 0; r++)
        {
            int d = -1, m = -1, l = -1, w = -1;
            for (int c = firstCol; c <= lastCol; c++)
            {
                string h = ws.Cell(r, c).GetString().Trim().ToUpperInvariant();
                if (h.Contains("DESCRIPTION") || h == "DESC") d = c;
                else if (h.Contains("MATERIAL")) m = c;
                else if (h.Contains("LENGTH") || h.Contains("AREA")) l = c;
                else if (h.Contains("WEIGHT")) w = c;
            }
            if (d > 0 && w > 0) { headerRow = r; cDesc = d; cMat = m; cLen = l; cWt = w; }
        }
        if (headerRow < 0) { headerRow = firstRow; cDesc = firstCol; cMat = firstCol + 1; cLen = firstCol + 2; cWt = firstCol + 3; }
        if (cMat < 0) cMat = cDesc + 1;
        if (cLen < 0) cLen = cMat + 1;
        if (cWt < 0) cWt = cLen + 1;

        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            string desc = ws.Cell(r, cDesc).GetString().Trim();
            string mat = ws.Cell(r, cMat).GetString().Trim();
            string len = CleanNumber(ws.Cell(r, cLen).GetString().Trim(), 3); // length/area: 3 dp
            string wt = CleanNumber(ws.Cell(r, cWt).GetString().Trim(), 2);   // weight: 2 dp

            bool isTotal = desc.Equals("Total", StringComparison.OrdinalIgnoreCase) ||
                           mat.Equals("Total", StringComparison.OrdinalIgnoreCase) ||
                           len.Equals("Total", StringComparison.OrdinalIgnoreCase) ||
                           (desc.Length == 0 && mat.Length == 0 && wt.Length > 0);
            if (isTotal)
            {
                if (wt.Length > 0) totalWeight = wt;
                continue;
            }
            if (desc.Length == 0 && mat.Length == 0 && len.Length == 0 && wt.Length == 0) continue;
            rows.Add(new[] { desc, mat, len, wt });
        }
        return (rows, totalWeight);
    }

    // Count support-list data rows on a sheet (data starts at row 3; a row counts if it has an
    // S.L. or a SUPPORT No.). Used to auto-pick the fullest batch sheet.
    private static int CountDataRows(IXLWorksheet ws)
    {
        var used = ws.RangeUsed();
        if (used == null) return 0;
        int last = used.LastRow().RowNumber(), count = 0;
        for (int r = 3; r <= last; r++)
            if (ws.Cell(r, 1).GetString().Trim().Length > 0 || ws.Cell(r, 2).GetString().Trim().Length > 0)
                count++;
        return count;
    }

    private static List<string[]> ReadExcel(string excelPath, string? sheetName)
    {
        using var wb = new XLWorkbook(excelPath);

        IXLWorksheet? ws = null;
        if (!string.IsNullOrWhiteSpace(sheetName))
        {
            var want = sheetName.Trim();
            var wantNoSpace = want.Replace(" ", "");
            ws = wb.Worksheets.FirstOrDefault(s => s.Name.Trim().Equals(want, StringComparison.OrdinalIgnoreCase))
              ?? wb.Worksheets.FirstOrDefault(s => s.Name.Replace(" ", "").Contains(wantNoSpace, StringComparison.OrdinalIgnoreCase));
        }
        // No sheet specified: a workbook can hold several batch sheets (BATCH-1/2/3). Pick the
        // support-list sheet with the MOST data rows so nothing is silently dropped (e.g. a
        // Batch-3 file uses BATCH-3's 568 rows, not BATCH-2's 563).
        ws ??= wb.Worksheets
            .Select(s => (Sheet: s, Count: CountDataRows(s)))
            .Where(x => x.Count > 0)
            .OrderByDescending(x => x.Count)
            .Select(x => x.Sheet)
            .FirstOrDefault();
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

    // Data font size (half-points): 18 = 9pt.
    private const string DataFontHalfPt = "18";

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

    private static void SetCellText(TableCell cell, string text, bool fitOneLine = false, bool center = false)
    {
        var para = cell.Elements<Paragraph>().FirstOrDefault();
        if (para == null)
        {
            para = new Paragraph();
            cell.Append(para);
        }

        if (center)
        {
            para.ParagraphProperties ??= new ParagraphProperties();
            para.ParagraphProperties.Justification = new Justification { Val = JustificationValues.Center };
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

            // Fill Rev / Date wherever the template shows them (page headers + book statistics).
            // Wildcards so it works for any template value (old "Rev. 00"/"24-07-2026",
            // new "Rev. A"/"21-04-2026", ...). The cover block is already set structurally.
            if (!string.IsNullOrWhiteSpace(date))
            {
                ReplaceEverywhere(document, "Date: [0-9]{1,2}-[0-9]{1,2}-[0-9]{4}", "Date: " + date, wildcards: true);
                ReplaceEverywhere(document, "Date: [0-9]{1,2}/[0-9]{1,2}/[0-9]{4}", "Date: " + date, wildcards: true);
            }
            if (!string.IsNullOrWhiteSpace(rev))
            {
                ReplaceEverywhere(document, "Rev. [0-9A-Za-z]{1,3}", "Rev. " + rev, wildcards: true);
                ReplaceEverywhere(document, "REV [0-9A-Za-z]{1,3}", "REV " + rev, wildcards: true);
                ReplaceEverywhere(document, "rev-[0-9A-Za-z]{1,3}", "rev-" + rev, wildcards: true);
            }

            // Replace the template's hardcoded sheet total ("Sheet X of N") with the real
            // count of booklet pages (cover..Appendix B) — NOT the appended drawings. A
            // wildcard handles any template value ("of 19", "of 7", ...).
            try
            {
                int total = (int)document.ComputeStatistics(2); // wdStatisticPages
                ReplaceEverywhere(document, " of [0-9]{1,3}", " of " + total, wildcards: true);
            }
            catch { }

            // Refresh the Table(s) of Contents so page numbers are correct and the
            // Appendix A/B entries are all listed, then update remaining fields.
            try { foreach (dynamic toc in document.TablesOfContents) toc.Update(); } catch { }
            try { document.Fields.Update(); } catch { }

            // Persist the filled + meta document so it can also be handed out as the Word
            // download (the .docx booklet body, without the appended drawings PDF).
            try { document.Save(); } catch { }

            document.ExportAsFixedFormat(pdfPath, 17); // wdExportFormatPDF
        }
        finally
        {
            try { document?.Close(false); } catch { }
            try { app.Quit(); } catch { }
        }
    }

    [SupportedOSPlatform("windows")]
    private static void ReplaceEverywhere(dynamic document, string find, string replace, bool wildcards = false)
    {
        ReplaceInRange(document.Content, find, replace, wildcards);
        foreach (dynamic section in document.Sections)
        {
            foreach (dynamic h in section.Headers) ReplaceInRange(h.Range, find, replace, wildcards);
            foreach (dynamic f in section.Footers) ReplaceInRange(f.Range, find, replace, wildcards);
        }
    }

    [SupportedOSPlatform("windows")]
    private static void ReplaceInRange(dynamic range, string find, string replace, bool wildcards = false)
    {
        try
        {
            dynamic f = range.Find;
            f.ClearFormatting();
            f.Replacement.ClearFormatting();
            // Execute(FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike,
            //   MatchAllWordForms, Forward, Wrap(1=wdFindContinue), Format, ReplaceWith, Replace(2=wdReplaceAll))
            f.Execute(find, false, false, wildcards, false, false, true, 1, false, replace, 2);
        }
        catch { /* a story without a Find range is fine to skip */ }
    }

    // ---------- Step 3: merge booklet PDF + drawings PDF ----------

    private static void MergePdfs(string bookletPdf, string drawingsPdf, string finalPdf)
    {
        // The drawings PDF is huge (hundreds of pages, 100+ MB); the booklet is ~20 pages.
        // Open the drawings doc once and insert only the booklet pages at the front, instead
        // of deep-copying every drawing page into a fresh document. That avoids rebuilding
        // the bulk of the file and is dramatically faster.
        try
        {
            using var outDoc = PdfReader.Open(drawingsPdf, PdfDocumentOpenMode.Modify);
            using var booklet = PdfReader.Open(bookletPdf, PdfDocumentOpenMode.Import);
            for (int i = booklet.PageCount - 1; i >= 0; i--)
                outDoc.InsertPage(0, booklet.Pages[i]);
            outDoc.Save(finalPdf);
        }
        catch
        {
            // Fallback: rebuild from scratch (slower, but tolerant of unusual PDFs).
            using var outDoc = new PdfDocument();
            AppendPages(outDoc, bookletPdf);
            AppendPages(outDoc, drawingsPdf);
            outDoc.Save(finalPdf);
        }
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
