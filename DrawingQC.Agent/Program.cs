using DrawingQC.Web;

// ---------------------------------------------------------------------------------------------
// DrawingQC Local Agent
//
// Runs on the engineer's Windows PC. The hosted web app (e.g. on Render) calls it directly on
// http://127.0.0.1:5081 so the features that need locally installed software keep working with a
// cloud-hosted UI:
//   GET  /status          what the agent can see (AutoCAD / Plant 3D product + open drawing, Word)
//   POST /sync-autocad    write the QC register table into the open AutoCAD drawing (COM)
//   POST /booklet         fill the Word template, export the PDF with Word, merge the drawings —
//                         the exact same BookletBuilder code the web app runs locally on Windows
//   GET  /booklet/download?format=pdf|word
// Only browsers on THIS machine can reach it (it binds to 127.0.0.1), and only pages from allowed
// origins (localhost, *.onrender.com, plus --origin / DRAWINGQC_AGENT_ORIGINS) pass the CORS check.
// ---------------------------------------------------------------------------------------------

int port = AgentInfo.DefaultPort;
var extraOrigins = new List<string>();
if (int.TryParse(Environment.GetEnvironmentVariable("DRAWINGQC_AGENT_PORT"), out var envPort)) port = envPort;
var envOrigins = Environment.GetEnvironmentVariable("DRAWINGQC_AGENT_ORIGINS");
if (!string.IsNullOrWhiteSpace(envOrigins))
    extraOrigins.AddRange(envOrigins.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries));
for (int i = 0; i + 1 < args.Length; i++)
{
    if (args[i] == "--port" && int.TryParse(args[i + 1], out var p)) port = p;
    if (args[i] == "--origin") extraOrigins.Add(args[i + 1].Trim());
}

var builder = WebApplication.CreateBuilder(new WebApplicationOptions { ContentRootPath = AppContext.BaseDirectory });
builder.Logging.ClearProviders();
builder.Logging.AddSimpleConsole(o => { o.SingleLine = true; o.TimestampFormat = "HH:mm:ss "; });
builder.Logging.SetMinimumLevel(LogLevel.Warning);
builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = null); // big drawings PDFs
builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(o => o.MultipartBodyLengthLimit = long.MaxValue);

var app = builder.Build();
app.Urls.Clear();
app.Urls.Add($"http://127.0.0.1:{port}");

bool OriginAllowed(string origin)
{
    if (!Uri.TryCreate(origin, UriKind.Absolute, out var u)) return false;
    if (u.Host is "localhost" or "127.0.0.1" or "[::1]") return true;
    if (u.Host.EndsWith(".onrender.com", StringComparison.OrdinalIgnoreCase)) return true;
    return extraOrigins.Any(o => string.Equals(o.TrimEnd('/'), origin.TrimEnd('/'), StringComparison.OrdinalIgnoreCase));
}

// CORS (+ Chrome "Private/Local Network Access" preflight) so an https:// page may call localhost.
app.Use(async (ctx, next) =>
{
    var origin = ctx.Request.Headers.Origin.ToString();
    bool hasOrigin = origin.Length > 0;
    bool allowed = hasOrigin && OriginAllowed(origin);
    if (allowed)
    {
        var h = ctx.Response.Headers;
        h["Access-Control-Allow-Origin"] = origin;
        h["Vary"] = "Origin";
        h["Access-Control-Allow-Methods"] = "GET, POST, OPTIONS";
        h["Access-Control-Allow-Headers"] = "Content-Type";
        h["Access-Control-Allow-Private-Network"] = "true";
        h["Access-Control-Max-Age"] = "600";
    }
    if (HttpMethods.IsOptions(ctx.Request.Method)) { ctx.Response.StatusCode = allowed || !hasOrigin ? 204 : 403; return; }
    if (hasOrigin && !allowed)
    {
        ctx.Response.StatusCode = 403;
        await ctx.Response.WriteAsJsonAsync(new { error = $"Origin not allowed: {origin}. Start the agent with --origin {origin} to allow it." });
        return;
    }
    await next();
});

void Log(string msg) => Console.WriteLine($"[{DateTime.Now:HH:mm:ss}] {msg}");

// ---------- status ----------

async Task<object> ProbeAsync()
{
    object autocad = new { running = false, product = (string?)null, document = (string?)null, error = "Windows only." };
    bool word = false;
    if (OperatingSystem.IsWindows())
    {
        var probe = AutoCadSync.StatusAsync();
        // A modal dialog inside AutoCAD can make COM calls block; don't let the page hang on it.
        if (await Task.WhenAny(probe, Task.Delay(3000)) == probe)
        {
            var s = probe.Result;
            autocad = new { running = s.Running, product = s.Product, document = s.Document, error = s.Error };
        }
        else autocad = new { running = true, product = "AutoCAD (busy)", document = (string?)null, error = "AutoCAD did not answer in time (a dialog may be open)." };
        word = Type.GetTypeFromProgID("Word.Application") != null;
    }
    return new
    {
        ok = true,
        agent = AgentInfo.Name,
        version = AgentInfo.Version,
        machine = Environment.MachineName,
        user = Environment.UserName,
        port,
        autocad,
        word = new { installed = word },
    };
}

app.MapGet("/status", async () => Results.Ok(await ProbeAsync()));

app.MapGet("/", () => Results.Content(Pages.Home, "text/html; charset=utf-8"));

// ---------- AutoCAD / Plant 3D: push the QC register into the open drawing ----------

app.MapPost("/sync-autocad", async (HttpRequest request) =>
{
    if (!OperatingSystem.IsWindows()) return Results.Problem("AutoCAD sync is only available on Windows.");
    SyncPayload? payload;
    try { payload = await request.ReadFromJsonAsync<SyncPayload>(); }
    catch (Exception ex) { return Results.BadRequest(new { error = "Invalid request body: " + ex.Message }); }
    if (payload?.Rows is null || payload.Rows.Count == 0)
        return Results.BadRequest(new { error = "No rows to sync. Run a QC check first." });

    var rows = payload.Rows
        .Select((r, i) => new QcRow(r.SrNo == 0 ? i + 1 : r.SrNo, r.FileName ?? "", r.Drawing1 ?? "", r.Drawing2 ?? "", r.Status ?? ""))
        .ToList();
    try
    {
        int count = await AutoCadSync.PushRegisterAsync(rows);
        Log($"Synced {count} QC rows into AutoCAD.");
        return Results.Ok(new { ok = true, count });
    }
    catch (Exception ex)
    {
        Log("AutoCAD sync failed: " + ex.Message);
        return Results.Problem(ex.Message);
    }
});

// ---------- Booklet: same inputs/outputs as the web app's /api/booklet ----------

string? lastBookletPdf = null;
string? lastBookletDocx = null;

app.MapPost("/booklet", async (HttpRequest request) =>
{
    if (!OperatingSystem.IsWindows()) return Results.Problem("Booklet generation is only available on Windows.");
    if (!request.HasFormContentType) return Results.BadRequest(new { error = "Expected a multipart form." });

    var form = await request.ReadFormAsync();
    var temps = new List<string>();
    try
    {
        string rev = form["rev"].ToString().Trim();
        string date = form["date"].ToString().Trim();
        string sheet = form["sheet"].ToString().Trim();
        string outputPath = form["outputPath"].ToString().Trim().Trim('"');
        string resolvedOutput = !string.IsNullOrWhiteSpace(outputPath)
            ? outputPath
            : Path.Combine(Path.GetTempPath(), $"booklet_out_{Guid.NewGuid():N}.pdf");

        string template = await ResolveInput(form, "template", ".docx", temps);
        string excel = await ResolveInput(form, "excel", ".xlsx", temps);
        string drawings = await ResolveInput(form, "drawings", ".pdf", temps);
        string bom = await ResolveInput(form, "bom", ".xlsx", temps); // optional

        Log($"Booklet: template={Path.GetFileName(template)} excel={Path.GetFileName(excel)} drawings={Path.GetFileName(drawings)}");
        var result = await Task.Run(() => BookletBuilder.Build(new BookletInputs
        {
            TemplatePath = template,
            ExcelPath = excel,
            DrawingsPath = drawings,
            BomPath = string.IsNullOrWhiteSpace(bom) ? null : bom,
            Sheet = string.IsNullOrWhiteSpace(sheet) ? null : sheet,
            Rev = rev,
            Date = date,
            Description = form["description"].ToString().Trim(),
            Prepared = form["prepared"].ToString().Trim(),
            Verified = form["verified"].ToString().Trim(),
            Approved = form["approved"].ToString().Trim(),
            OutputPath = resolvedOutput,
        }));

        lastBookletPdf = result.PdfPath;
        lastBookletDocx = result.DocxPath;
        var info = new FileInfo(result.PdfPath);
        Log($"Booklet done: {Path.GetFileName(result.PdfPath)} ({info.Length / 1024.0 / 1024.0:0.0} MB)");
        return Results.Ok(new
        {
            ok = true,
            outputPath = result.PdfPath,
            fileName = Path.GetFileName(result.PdfPath),
            docxFileName = Path.GetFileName(result.DocxPath),
            savedToTemp = string.IsNullOrWhiteSpace(outputPath),
            sizeMB = Math.Round(info.Length / 1024.0 / 1024.0, 1),
        });
    }
    catch (Exception ex)
    {
        Log("Booklet failed: " + ex.Message);
        return Results.Problem(ex.Message);
    }
    finally
    {
        foreach (var t in temps) { try { File.Delete(t); } catch { } }
    }
});

app.MapGet("/booklet/download", (string? format) =>
{
    bool word = string.Equals(format, "word", StringComparison.OrdinalIgnoreCase)
             || string.Equals(format, "docx", StringComparison.OrdinalIgnoreCase);
    string? path = word ? lastBookletDocx : lastBookletPdf;
    if (path is null || !File.Exists(path)) return Results.NotFound();
    string ctype = word
        ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        : "application/pdf";
    return Results.File(path, ctype, Path.GetFileName(path));
});

// ---------- run ----------

Console.Title = $"{AgentInfo.Name} {AgentInfo.Version}";
Console.WriteLine($"{AgentInfo.Name} {AgentInfo.Version}");
Console.WriteLine($"Listening on http://127.0.0.1:{port}   (keep this window open; close it to stop)");
Console.WriteLine("Allowed sites: localhost, *.onrender.com" + (extraOrigins.Count > 0 ? ", " + string.Join(", ", extraOrigins) : ""));
if (!OperatingSystem.IsWindows())
    Console.WriteLine("WARNING: not running on Windows — AutoCAD and Word features will be unavailable.");
else
{
    var s = await AutoCadSync.StatusAsync();
    Console.WriteLine(s.Running ? $"AutoCAD: {s.Product}" + (s.Document != null ? $" — {s.Document}" : "") : "AutoCAD: not running (open it before syncing)");
    Console.WriteLine(Type.GetTypeFromProgID("Word.Application") != null ? "Word: installed" : "Word: NOT found — booklet generation will fail");
}
Console.WriteLine();

try
{
    app.Run();
}
catch (Exception ex)
{
    Console.Error.WriteLine();
    Console.Error.WriteLine("The agent could not start: " + ex.Message);
    Console.Error.WriteLine($"Is another copy already running on port {port}? Start with --port <other> to change it.");
    if (!Console.IsInputRedirected) { Console.Error.WriteLine("Press any key to close."); Console.ReadKey(true); }
    return 1;
}
return 0;

// An input may arrive as an uploaded file (<name>File) or as a local path (<name>Path).
static async Task<string> ResolveInput(IFormCollection form, string name, string ext, List<string> temps)
{
    var file = form.Files.GetFile(name + "File");
    if (file is not null && file.Length > 0)
    {
        string tmp = Path.Combine(Path.GetTempPath(), $"booklet_{name}_{Guid.NewGuid():N}{ext}");
        await using var fs = File.Create(tmp);
        await file.CopyToAsync(fs);
        temps.Add(tmp);
        return tmp;
    }
    return form[name + "Path"].ToString().Trim().Trim('"');
}

// Tiny status page for people who open the agent URL in a browser.
static class Pages
{
    public const string Home = """
<!doctype html><html><head><meta charset="utf-8"><title>DrawingQC Local Agent</title>
<style>body{font:15px/1.5 system-ui,Segoe UI,sans-serif;max-width:560px;margin:60px auto;color:#131a2b;padding:0 16px}
h1{font-size:20px}code{background:#eef1f6;padding:2px 6px;border-radius:5px}.ok{color:#1b6b3a}.bad{color:#9c0006}</style></head>
<body><h1>DrawingQC Local Agent is running</h1>
<p>The Support Automation website talks to this agent to use the AutoCAD / Plant 3D and Microsoft Word installed on this PC.</p>
<ul id="st"><li>Checking…</li></ul>
<p>Leave the agent window open while you work. Close it to stop.</p>
<script>fetch('/status').then(r=>r.json()).then(s=>{const a=s.autocad||{};document.getElementById('st').innerHTML=
'<li>Agent <b>'+s.version+'</b> on <code>'+location.host+'</code> ('+s.machine+')</li>'+
'<li class="'+(a.running?'ok':'bad')+'">AutoCAD: '+(a.running?(a.product||'running')+(a.document?' — '+a.document:''):'not running')+'</li>'+
'<li class="'+(s.word&&s.word.installed?'ok':'bad')+'">Microsoft Word: '+(s.word&&s.word.installed?'installed':'not found')+'</li>';}).catch(()=>{});</script>
</body></html>
""";
}
