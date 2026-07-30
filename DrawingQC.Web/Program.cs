using System.Collections.Concurrent;
using DrawingQC.UI; // QcEngine + PdfResult live here (no WinForms dependency)

var builder = WebApplication.CreateBuilder(args);

// Allow large zip uploads. Two separate limits both need lifting:
//  - Kestrel's request body size cap (default 30 MB)
//  - the multipart form body length limit (default ~128 MB)
builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = null); // unlimited
builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(o =>
{
    o.MultipartBodyLengthLimit = long.MaxValue;
});

var app = builder.Build();

// Locally: keep http://localhost:3000 exactly as before.
// When a hosting platform (e.g. Render) provides a PORT, bind to that instead.
app.Urls.Clear();
var hostPort = Environment.GetEnvironmentVariable("PORT");
app.Urls.Add(string.IsNullOrEmpty(hostPort)
    ? "http://localhost:3000"
    : $"http://0.0.0.0:{hostPort}");

app.UseDefaultFiles();
app.UseStaticFiles();

// Generated Excel reports, kept in memory keyed by a token so the browser can download them.
var reports = new ConcurrentDictionary<string, (byte[] Bytes, string FileName)>();

app.MapPost("/api/analyze", async (HttpRequest request) =>
{
    if (!request.HasFormContentType)
        return Results.BadRequest(new { error = "Expected a multipart form upload." });

    var form = await request.ReadFormAsync();
    var file = form.Files.GetFile("file");
    if (file is null || file.Length == 0)
        return Results.BadRequest(new { error = "No zip file was uploaded." });

    if (!file.FileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        return Results.BadRequest(new { error = "Please upload a .zip file of drawing PDFs." });

    // QcEngine works off a file path, so buffer the upload to a temp file.
    var tempZip = Path.Combine(Path.GetTempPath(), $"dqc_{Guid.NewGuid():N}.zip");
    try
    {
        await using (var fs = File.Create(tempZip))
            await file.CopyToAsync(fs);

        List<PdfResult> results = await Task.Run(() => QcEngine.Analyze(tempZip));

        // Build the Excel report in memory and stash it for download.
        var tempXlsx = Path.Combine(Path.GetTempPath(), $"dqc_{Guid.NewGuid():N}.xlsx");
        byte[] xlsx;
        try
        {
            QcEngine.WriteReport(results, tempXlsx);
            xlsx = await File.ReadAllBytesAsync(tempXlsx);
        }
        finally
        {
            if (File.Exists(tempXlsx)) File.Delete(tempXlsx);
        }

        var token = Guid.NewGuid().ToString("N");
        var reportName = Path.GetFileNameWithoutExtension(file.FileName) + " - QC Report.xlsx";
        reports[token] = (xlsx, reportName);

        var rows = results.Select((r, i) => new
        {
            srNo = i + 1,
            fileName = r.FileName,
            drawing1 = string.IsNullOrWhiteSpace(r.Drawing1) ? "-" : r.Drawing1,
            drawing2 = string.IsNullOrWhiteSpace(r.Drawing2) ? "-" : r.Drawing2,
            status = r.FinalStatus,
            error = r.Error,
        });

        var summary = new
        {
            total = results.Count,
            matched = results.Count(x => x.FinalStatus == "Matched"),
            unmatched = results.Count(x => x.FinalStatus == "Unmatched"),
            duplicate = results.Count(x => x.FinalStatus == "Duplicate"),
        };

        return Results.Ok(new { reportToken = token, reportName, summary, rows });
    }
    catch (Exception ex)
    {
        return Results.Problem($"Failed to analyze zip: {ex.Message}");
    }
    finally
    {
        if (File.Exists(tempZip)) File.Delete(tempZip);
    }
});

app.MapGet("/api/report/{token}", (string token) =>
{
    if (!reports.TryGetValue(token, out var report))
        return Results.NotFound();

    return Results.File(
        report.Bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        report.FileName);
});

// Push the QC register into the AutoCAD instance running on this PC (live COM sync).
app.MapPost("/api/sync-autocad", async (HttpRequest request) =>
{
    if (!OperatingSystem.IsWindows())
        return Results.Problem("AutoCAD sync is only available on Windows.");

    DrawingQC.Web.SyncPayload? payload;
    try
    {
        payload = await request.ReadFromJsonAsync<DrawingQC.Web.SyncPayload>();
    }
    catch (Exception ex)
    {
        return Results.BadRequest(new { error = "Invalid request body: " + ex.Message });
    }

    if (payload?.Rows is null || payload.Rows.Count == 0)
        return Results.BadRequest(new { error = "No rows to sync. Run a QC check first." });

    var rows = payload.Rows
        .Select((r, i) => new DrawingQC.Web.QcRow(
            r.SrNo == 0 ? i + 1 : r.SrNo,
            r.FileName ?? string.Empty,
            r.Drawing1 ?? string.Empty,
            r.Drawing2 ?? string.Empty,
            r.Status ?? string.Empty))
        .ToList();

    try
    {
        int count = await DrawingQC.Web.AutoCadSync.PushRegisterAsync(rows);
        return Results.Ok(new { ok = true, count });
    }
    catch (Exception ex)
    {
        // Surface a clean message (e.g. "no running AutoCAD") to the UI.
        return Results.Problem(ex.Message);
    }
});

app.Run();