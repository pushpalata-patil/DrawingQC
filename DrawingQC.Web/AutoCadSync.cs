using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace DrawingQC.Web;

/// <summary>One row of the QC register that gets written into AutoCAD.</summary>
public sealed record QcRow(int SrNo, string FileName, string Drawing1, string Drawing2, string Status);

/// <summary>JSON shapes for the /api/sync-autocad request body.</summary>
public sealed class SyncPayload { public List<SyncRow>? Rows { get; set; } }
public sealed class SyncRow
{
    public int SrNo { get; set; }
    public string? FileName { get; set; }
    public string? Drawing1 { get; set; }
    public string? Drawing2 { get; set; }
    public string? Status { get; set; }
}

/// <summary>
/// Pushes the QC register into a running AutoCAD on this machine via late-bound COM.
/// No AutoCAD interop assemblies are referenced — we grab the running application from
/// the COM Running Object Table and drive it through IDispatch (dynamic).
/// </summary>
[SupportedOSPlatform("windows")]
public static class AutoCadSync
{
    [DllImport("ole32.dll", PreserveSig = false)]
    private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object obj);

    /// <summary>Runs the COM work on a dedicated STA thread (AutoCAD's COM server is STA).</summary>
    public static Task<int> PushRegisterAsync(IReadOnlyList<QcRow> rows)
    {
        var tcs = new TaskCompletionSource<int>();
        var thread = new Thread(() =>
        {
            try { tcs.SetResult(PushRegister(rows)); }
            catch (Exception ex) { tcs.SetException(ex); }
        })
        { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        return tcs.Task;
    }

    private static int PushRegister(IReadOnlyList<QcRow> rows)
    {
        dynamic app = GetRunningAutoCad();
        try { app.Visible = true; } catch { /* non-fatal */ }

        dynamic doc;
        try { doc = app.ActiveDocument; }
        catch { throw new InvalidOperationException("AutoCAD is running but has no drawing open. Open or create a drawing, then sync again."); }

        dynamic modelSpace = doc.ModelSpace;

        int dataRows = rows.Count;
        int nRows = dataRows + 2;   // title row + header row + data rows
        const int nCols = 5;
        const double rowHeight = 6.0;
        const double colWidth = 40.0;

        double[] insertion = { 0.0, 0.0, 0.0 };
        dynamic table = modelSpace.AddTable(insertion, nRows, nCols, rowHeight, colWidth);

        // Column widths in drawing units (best-effort — ignore if a version rejects it).
        try
        {
            table.SetColumnWidth(0, 15.0);   // Sr.No
            table.SetColumnWidth(1, 95.0);   // Support PDF
            table.SetColumnWidth(2, 55.0);   // 1st drawing
            table.SetColumnWidth(3, 55.0);   // 2nd drawing
            table.SetColumnWidth(4, 35.0);   // Status
        }
        catch { /* keep default widths */ }

        // Title (row 0) + column headers (row 1).
        table.SetText(0, 0, "QC Register");
        string[] headers = { "Sr.No", "Support PDF", "1st drawing", "2nd drawing", "Status" };
        for (int c = 0; c < nCols; c++) table.SetText(1, c, headers[c]);

        // Data rows (row 2 onward).
        for (int i = 0; i < dataRows; i++)
        {
            int r = i + 2;
            QcRow row = rows[i];
            table.SetText(r, 0, row.SrNo.ToString());
            table.SetText(r, 1, row.FileName ?? "");
            table.SetText(r, 2, string.IsNullOrWhiteSpace(row.Drawing1) ? "-" : row.Drawing1);
            table.SetText(r, 3, string.IsNullOrWhiteSpace(row.Drawing2) ? "-" : row.Drawing2);
            table.SetText(r, 4, row.Status ?? "");
        }

        try { app.ZoomExtents(); } catch { /* view refresh is non-fatal */ }
        try { app.Update(); } catch { }

        return dataRows;
    }

    private static object GetRunningAutoCad()
    {
        // Version-independent ProgID first, then the versions installed on this machine
        // (R26.0 = AutoCAD 2027, R24.3 = AutoCAD 2024).
        string[] progIds = { "AutoCAD.Application", "AutoCAD.Application.26", "AutoCAD.Application.24" };
        foreach (string progId in progIds)
        {
            try
            {
                CLSIDFromProgID(progId, out Guid clsid);
                GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
                if (obj is not null) return obj;
            }
            catch { /* try the next ProgID */ }
        }
        throw new InvalidOperationException(
            "No running AutoCAD found. Open AutoCAD (with a drawing) on this PC, then click Sync again.");
    }
}
