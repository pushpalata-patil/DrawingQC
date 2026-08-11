using System.IO.Compression;
using ClosedXML.Excel;

namespace DrawingQC.Web;

/// <summary>One expected drawing and whether a matching delivered file was found.</summary>
public sealed class TagDeliveryRow
{
    public int SrNo { get; set; }
    public string DrawingNo { get; set; } = "";
    public string Description { get; set; } = "";
    public string Status { get; set; } = "";        // Delivered | Pending
    public string DeliveryDate { get; set; } = "";   // date of the matched delivered file
    public string MatchedFile { get; set; } = "";
}

public sealed class TagDeliveryResult
{
    public int Total { get; set; }
    public int Delivered { get; set; }
    public int Pending { get; set; }
    public int DoneFiles { get; set; }
    public string SheetUsed { get; set; } = "";
    public string Column { get; set; } = "";
    public List<TagDeliveryRow> Rows { get; set; } = new();
    public byte[] Excel { get; set; } = Array.Empty<byte>();
}

/// <summary>
/// KBR "Tagwise Delivery Report": take a client Excel (the expected drawing list) and a zip
/// of the delivered/done .dwg/.pdf files. Each Excel drawing number is marked Delivered if a
/// delivered file's name contains that number, else Pending. Produces a colour-coded Excel.
/// </summary>
public static class TagDeliveryReport
{
    private static readonly string[] DoneExtensions = { ".dwg", ".pdf" };

    public static TagDeliveryResult Build(string clientExcelPath, string zipPath)
    {
        // 1) Delivered file names from the zip (.dwg / .pdf), normalized for matching.
        var delivered = new List<(string file, string norm, DateTime date)>();
        using (var archive = ZipFile.OpenRead(zipPath))
            foreach (var e in archive.Entries)
            {
                if (string.IsNullOrEmpty(e.Name)) continue; // directory entry
                if (DoneExtensions.Any(x => e.Name.EndsWith(x, StringComparison.OrdinalIgnoreCase)))
                    delivered.Add((e.Name, Norm(Path.GetFileNameWithoutExtension(e.Name)), e.LastWriteTime.DateTime));
            }

        // 2) Client Excel: find the drawing-number column.
        using var wb = new XLWorkbook(clientExcelPath);
        var ws = wb.Worksheets
            .Select(s => (s, rows: s.RangeUsed()?.RowCount() ?? 0))
            .OrderByDescending(x => x.rows).Select(x => x.s).FirstOrDefault() ?? wb.Worksheets.First();
        var used = ws.RangeUsed() ?? throw new InvalidOperationException("The client Excel sheet is empty.");
        int firstRow = used.FirstRow().RowNumber(), lastRow = used.LastRow().RowNumber();
        int firstCol = used.FirstColumn().ColumnNumber(), lastCol = used.LastColumn().ColumnNumber();

        // Header row: first row (within the top 15) that has at least two non-empty cells.
        int headerRow = firstRow;
        for (int r = firstRow; r <= Math.Min(firstRow + 14, lastRow); r++)
        {
            int nonEmpty = 0;
            for (int c = firstCol; c <= lastCol; c++)
                if (ws.Cell(r, c).GetString().Trim().Length > 0) nonEmpty++;
            if (nonEmpty >= 2) { headerRow = r; break; }
        }

        // Pick the drawing-number column by scoring each column's VALUES (not just the header):
        // codes with letters+digits/hyphens or long numbers score high; short serial integers
        // (1,2,3…) score low. A matching header adds a bonus; a serial-ish header is penalised.
        int dwgCol = firstCol, descCol = -1;
        string colHeader = "";
        double best = double.MinValue;
        int sampleEnd = Math.Min(headerRow + 60, lastRow);
        for (int c = firstCol; c <= lastCol; c++)
        {
            string header = ws.Cell(headerRow, c).GetString().Trim();
            string hl = header.ToLowerInvariant();
            int n = 0, dwgLike = 0, serialLike = 0;
            for (int r = headerRow + 1; r <= sampleEnd; r++)
            {
                var v = ws.Cell(r, c).GetString().Trim();
                if (v.Length == 0) continue;
                n++;
                bool pureInt = long.TryParse(v.Replace(",", ""), out _);
                bool code = v.Contains('-') || v.Contains('/') || (v.Any(char.IsLetter) && v.Any(char.IsDigit));
                if (code && !pureInt) dwgLike++;
                else if (pureInt && v.Length >= 5) dwgLike++;   // long numeric drawing numbers
                else if (pureInt) serialLike++;                 // short 1-4 digit = serial index
            }
            double score = n > 0 ? (10.0 * dwgLike / n) - (3.0 * serialLike / n) : 0;
            if (hl.Contains("drawing") || hl.Contains("dwg") || hl.Contains("drg") || hl.Contains("document") ||
                hl.Contains("doc no") || hl.Contains("p&id") || hl.Contains("pid") || hl.Contains("sheet no") ||
                hl.Contains("tag")) score += 5;
            if (hl is "sr" or "sr no" or "sr.no" or "s.no" or "no" or "serial" || hl.Contains("sl no")) score -= 5;
            if (score > best) { best = score; dwgCol = c; colHeader = header; }

            if (descCol < 0 && (hl.Contains("desc") || hl.Contains("title") || hl.Contains("item"))) descCol = c;
        }
        if (descCol == dwgCol) descCol = -1;

        // 3) Match each expected drawing against the delivered file names.
        var result = new TagDeliveryResult { SheetUsed = ws.Name, Column = colHeader, DoneFiles = delivered.Count };
        int sr = 1;
        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            string dwg = ws.Cell(r, dwgCol).GetString().Trim();
            if (dwg.Length == 0) continue;
            string desc = descCol > 0 ? ws.Cell(r, descCol).GetString().Trim() : "";

            string key = Norm(dwg);
            var hit = key.Length >= 3 ? delivered.FirstOrDefault(f => f.norm.Contains(key)) : default;
            string matched = hit.file ?? "";
            string date = hit.file != null ? hit.date.ToString("dd-MM-yyyy") : "";

            bool isDelivered = matched.Length > 0;
            result.Rows.Add(new TagDeliveryRow
            {
                SrNo = sr++,
                DrawingNo = dwg,
                Description = desc,
                Status = isDelivered ? "Delivered" : "Pending",
                DeliveryDate = date,
                MatchedFile = matched,
            });
            if (isDelivered) result.Delivered++; else result.Pending++;
        }
        result.Total = result.Rows.Count;
        result.Excel = WriteExcel(result);
        return result;
    }

    // Uppercase and strip whitespace so "265 S2N ACD 5003" matches "265-S2N-ACD-5003_R0".
    private static string Norm(string s) =>
        new string((s ?? "").Where(ch => !char.IsWhiteSpace(ch)).ToArray()).ToUpperInvariant();

    private static byte[] WriteExcel(TagDeliveryResult r)
    {
        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Delivery Status");

        ws.Cell(1, 1).Value = "Tagwise Delivery Report";
        ws.Range(1, 1, 1, 6).Merge().Style.Font.SetBold().Font.FontSize = 14;

        string[] headers = { "S.No", "Drawing No.", "Description", "Status", "Delivery Date", "Delivered File" };
        for (int c = 0; c < headers.Length; c++)
        {
            var cell = ws.Cell(3, c + 1);
            cell.Value = headers[c];
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontColor = XLColor.White;
            cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#374757");
        }

        var green = XLColor.FromHtml("#C6EFCE"); var greenF = XLColor.FromHtml("#006100");
        var red = XLColor.FromHtml("#FFC7CE"); var redF = XLColor.FromHtml("#9C0006");

        int row = 4;
        foreach (var t in r.Rows)
        {
            ws.Cell(row, 1).Value = t.SrNo;
            ws.Cell(row, 2).Value = t.DrawingNo;
            ws.Cell(row, 3).Value = t.Description;
            ws.Cell(row, 4).Value = t.Status;
            ws.Cell(row, 5).Value = t.DeliveryDate;
            ws.Cell(row, 6).Value = t.MatchedFile;
            var rng = ws.Range(row, 1, row, 6);
            if (t.Status == "Delivered") { rng.Style.Fill.BackgroundColor = green; rng.Style.Font.FontColor = greenF; }
            else { rng.Style.Fill.BackgroundColor = red; rng.Style.Font.FontColor = redF; }
            row++;
        }
        ws.SheetView.FreezeRows(3);
        ws.Columns().AdjustToContents();

        var sum = wb.Worksheets.Add("Summary");
        sum.Cell(1, 1).Value = "Delivery Summary";
        sum.Cell(1, 1).Style.Font.SetBold().Font.FontSize = 14;
        sum.Cell(3, 1).Value = "Total items"; sum.Cell(3, 2).Value = r.Total;
        sum.Cell(4, 1).Value = "Delivered"; sum.Cell(4, 2).Value = r.Delivered; sum.Range(4, 1, 4, 2).Style.Fill.BackgroundColor = green;
        sum.Cell(5, 1).Value = "Pending"; sum.Cell(5, 2).Value = r.Pending; sum.Range(5, 1, 5, 2).Style.Fill.BackgroundColor = red;
        sum.Cell(6, 1).Value = "Done files in zip"; sum.Cell(6, 2).Value = r.DoneFiles;
        sum.Columns().AdjustToContents();

        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        return ms.ToArray();
    }
}
