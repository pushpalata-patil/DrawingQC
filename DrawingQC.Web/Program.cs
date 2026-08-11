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

// ---------- Authentication: accounts, login/logout, profile ----------
app.MapPost("/api/auth/register", (HttpContext ctx, RegisterDto dto) =>
{
    var (ok, err, user) = DrawingQC.Web.Auth.Register(dto.username, dto.email, dto.password, dto.name, dto.role, dto.securityQuestion, dto.securityAnswer);
    if (!ok) return Results.BadRequest(new { error = err });
    SetSession(ctx, user!.Id, remember: true);
    return Results.Ok(new { user = DrawingQC.Web.Auth.Public(user) });
});

// Self-service password reset: enter email + a new password.
app.MapPost("/api/auth/reset", (ResetDto dto) =>
{
    var (ok, err) = DrawingQC.Web.Auth.ResetPassword(dto.email, dto.password);
    return ok ? Results.Ok(new { ok = true }) : Results.BadRequest(new { error = err });
});

app.MapPost("/api/auth/security", (HttpContext ctx, SecurityDto dto) =>
{
    var user = CurrentUser(ctx.Request);
    if (user == null) return Results.Json(new { error = "Not signed in." }, statusCode: 401);
    return DrawingQC.Web.Auth.SetSecurityQuestion(user.Id, dto.question, dto.answer)
        ? Results.Ok(new { ok = true })
        : Results.BadRequest(new { error = "Please provide both a question and an answer." });
});

app.MapPost("/api/auth/login", (HttpContext ctx, LoginDto dto) =>
{
    var user = DrawingQC.Web.Auth.Validate(dto.login, dto.password);
    if (user == null) return Results.Json(new { error = "Invalid username or password." }, statusCode: 401);
    SetSession(ctx, user.Id, dto.remember);
    return Results.Ok(new { user = DrawingQC.Web.Auth.Public(user) });
});

app.MapPost("/api/auth/logout", (HttpContext ctx) => { ClearSession(ctx); return Results.Ok(new { ok = true }); });

app.MapGet("/api/auth/me", (HttpContext ctx) =>
{
    var user = CurrentUser(ctx.Request);
    return user == null
        ? Results.Json(new { error = "Not signed in." }, statusCode: 401)
        : Results.Ok(new { user = DrawingQC.Web.Auth.Public(user) });
});

app.MapPost("/api/auth/change-password", (HttpContext ctx, ChangePwDto dto) =>
{
    var user = CurrentUser(ctx.Request);
    if (user == null) return Results.Json(new { error = "Not signed in." }, statusCode: 401);
    return DrawingQC.Web.Auth.ChangePassword(user.Id, dto.current, dto.next)
        ? Results.Ok(new { ok = true })
        : Results.BadRequest(new { error = "Current password is incorrect, or the new password is too short (min 6)." });
});

app.MapPost("/api/auth/profile", (HttpContext ctx, ProfileDto dto) =>
{
    var user = CurrentUser(ctx.Request);
    if (user == null) return Results.Json(new { error = "Not signed in." }, statusCode: 401);
    var updated = DrawingQC.Web.Auth.UpdateProfile(user.Id, dto.name, dto.email, dto.role, dto.avatar);
    return updated == null
        ? Results.BadRequest(new { error = "Could not update profile (email may already be in use)." })
        : Results.Ok(new { user = DrawingQC.Web.Auth.Public(updated) });
});

// Public config so the login page knows whether to offer "Create an account".
app.MapGet("/api/auth/config", () => Results.Ok(new
{
    registrationOpen = DrawingQC.Web.Auth.GetSettings().RegistrationOpen,
    hasUsers = DrawingQC.Web.Auth.UserCount() > 0,
}));

// ---------- Admin: user management (admins only) ----------
app.MapGet("/api/admin/users", (HttpContext ctx) =>
{
    if (!DrawingQC.Web.Auth.IsAdmin(CurrentUser(ctx.Request))) return Results.Json(new { error = "Admins only." }, statusCode: 403);
    return Results.Ok(new { users = DrawingQC.Web.Auth.ListUsers(), registrationOpen = DrawingQC.Web.Auth.GetSettings().RegistrationOpen });
});

app.MapPost("/api/admin/role", (HttpContext ctx, AdminRoleDto dto) =>
{
    if (!DrawingQC.Web.Auth.IsAdmin(CurrentUser(ctx.Request))) return Results.Json(new { error = "Admins only." }, statusCode: 403);
    var (ok, err) = DrawingQC.Web.Auth.AdminSetRole(dto.id ?? "", dto.role);
    return ok ? Results.Ok(new { ok = true }) : Results.BadRequest(new { error = err });
});

app.MapPost("/api/admin/reset", (HttpContext ctx, AdminResetDto dto) =>
{
    if (!DrawingQC.Web.Auth.IsAdmin(CurrentUser(ctx.Request))) return Results.Json(new { error = "Admins only." }, statusCode: 403);
    var (ok, err) = DrawingQC.Web.Auth.AdminResetPassword(dto.id ?? "", dto.password);
    return ok ? Results.Ok(new { ok = true }) : Results.BadRequest(new { error = err });
});

app.MapPost("/api/admin/delete", (HttpContext ctx, AdminIdDto dto) =>
{
    var me = CurrentUser(ctx.Request);
    if (!DrawingQC.Web.Auth.IsAdmin(me)) return Results.Json(new { error = "Admins only." }, statusCode: 403);
    var (ok, err) = DrawingQC.Web.Auth.AdminDelete(me!.Id, dto.id ?? "");
    return ok ? Results.Ok(new { ok = true }) : Results.BadRequest(new { error = err });
});

app.MapPost("/api/admin/registration", (HttpContext ctx, RegToggleDto dto) =>
{
    if (!DrawingQC.Web.Auth.IsAdmin(CurrentUser(ctx.Request))) return Results.Json(new { error = "Admins only." }, statusCode: 403);
    DrawingQC.Web.Auth.SetRegistrationOpen(dto.open);
    return Results.Ok(new { ok = true, registrationOpen = dto.open });
});

// Generated Excel reports, kept in memory keyed by a token so the browser can download them.
var reports = new ConcurrentDictionary<string, (byte[] Bytes, string FileName)>();

app.MapPost("/api/analyze", async (HttpRequest request) =>
{
    if (CurrentUser(request) == null)
        return Results.Json(new { error = "Please sign in first." }, statusCode: 401);
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
    if (CurrentUser(request) == null)
        return Results.Json(new { error = "Please sign in first." }, statusCode: 401);
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

// ---------- Booklet tool (QATAR): fill Word template + append drawings -> merged PDF ----------
string? lastBookletPdf = null;
string? lastBookletDocx = null;

app.MapPost("/api/booklet", async (HttpRequest request) =>
{
    if (CurrentUser(request) == null)
        return Results.Json(new { error = "Please sign in first." }, statusCode: 401);
    if (!OperatingSystem.IsWindows())
        return Results.Problem("Booklet generation is only available on the local Windows app (it needs Microsoft Word installed).");
    if (!request.HasFormContentType)
        return Results.BadRequest(new { error = "Expected a multipart form." });

    var form = await request.ReadFormAsync();
    var temps = new List<string>();
    try
    {
        string rev = form["rev"].ToString().Trim();
        string date = form["date"].ToString().Trim();
        string description = form["description"].ToString().Trim();
        string prepared = form["prepared"].ToString().Trim();
        string verified = form["verified"].ToString().Trim();
        string approved = form["approved"].ToString().Trim();
        string sheet = form["sheet"].ToString().Trim();
        string outputPath = form["outputPath"].ToString().Trim().Trim('"');
        // Where the merged PDF is written: an explicit path wins; otherwise a temp file that
        // the user grabs via the "Download booklet" button.
        string? resolvedOutput = !string.IsNullOrWhiteSpace(outputPath)
            ? outputPath
            : Path.Combine(Path.GetTempPath(), $"booklet_out_{Guid.NewGuid():N}.pdf");

        string template = await ResolveBookletInput(form, "template", ".docx", temps);
        string excel = await ResolveBookletInput(form, "excel", ".xlsx", temps);
        string drawings = await ResolveBookletInput(form, "drawings", ".pdf", temps);
        string bom = await ResolveBookletInput(form, "bom", ".xlsx", temps); // optional

        var result = await Task.Run(() => DrawingQC.Web.BookletBuilder.Build(new DrawingQC.Web.BookletInputs
        {
            TemplatePath = template,
            ExcelPath = excel,
            DrawingsPath = drawings,
            BomPath = string.IsNullOrWhiteSpace(bom) ? null : bom,
            Sheet = string.IsNullOrWhiteSpace(sheet) ? null : sheet,
            Rev = rev,
            Date = date,
            Description = description,
            Prepared = prepared,
            Verified = verified,
            Approved = approved,
            OutputPath = resolvedOutput,
        }));

        lastBookletPdf = result.PdfPath;
        lastBookletDocx = result.DocxPath;
        var info = new FileInfo(result.PdfPath);
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
        return Results.Problem(ex.Message);
    }
    finally
    {
        foreach (var t in temps) { try { File.Delete(t); } catch { } }
    }
});

app.MapGet("/api/booklet/download", (string? format) =>
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

// ---------- KBR: Tagwise Delivery Report ----------
app.MapPost("/api/kbr/tagreport", async (HttpRequest request) =>
{
    if (CurrentUser(request) == null)
        return Results.Json(new { error = "Please sign in first." }, statusCode: 401);
    if (!request.HasFormContentType)
        return Results.BadRequest(new { error = "Expected a form upload." });

    var form = await request.ReadFormAsync();
    var temps = new List<string>();
    try
    {
        string excel = await ResolveBookletInput(form, "excel", ".xlsx", temps);   // client list
        string zip = await ResolveBookletInput(form, "zip", ".zip", temps);        // done files
        if (string.IsNullOrWhiteSpace(excel) || !File.Exists(excel))
            return Results.BadRequest(new { error = "Please provide the client Excel (upload a file or paste a valid path)." });
        if (string.IsNullOrWhiteSpace(zip) || !File.Exists(zip))
            return Results.BadRequest(new { error = "Please provide the .zip of delivered .dwg/.pdf files." });

        var result = await Task.Run(() => DrawingQC.Web.TagDeliveryReport.Build(excel, zip));
        var token = Guid.NewGuid().ToString("N");
        reports[token] = (result.Excel, "Tagwise Delivery Report.xlsx");
        return Results.Ok(new
        {
            ok = true,
            reportToken = token,
            summary = new { total = result.Total, delivered = result.Delivered, pending = result.Pending, doneFiles = result.DoneFiles, sheet = result.SheetUsed, column = result.Column },
            rows = result.Rows,
        });
    }
    catch (Exception ex)
    {
        return Results.Problem("Failed to build report: " + ex.Message);
    }
    finally
    {
        foreach (var t in temps) { try { File.Delete(t); } catch { } }
    }
});

app.Run();

// An input may arrive as an uploaded file (<name>File) or as a local path (<name>Path).
static async Task<string> ResolveBookletInput(IFormCollection form, string name, string ext, List<string> temps)
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

// ---------- Auth session helpers ----------
static DrawingQC.Web.UserAccount? CurrentUser(HttpRequest req)
{
    var uid = DrawingQC.Web.Auth.ValidateToken(req.Cookies[DrawingQC.Web.Auth.CookieName]);
    return uid == null ? null : DrawingQC.Web.Auth.FindById(uid);
}

static void SetSession(HttpContext ctx, string userId, bool remember)
{
    var exp = DateTime.UtcNow.AddDays(remember ? 30 : 1);
    var token = DrawingQC.Web.Auth.CreateToken(userId, exp);
    var opts = new CookieOptions { HttpOnly = true, SameSite = SameSiteMode.Lax, Path = "/" };
    if (remember) opts.Expires = new DateTimeOffset(exp); // persistent "remember me" cookie
    ctx.Response.Cookies.Append(DrawingQC.Web.Auth.CookieName, token, opts);
}

static void ClearSession(HttpContext ctx) =>
    ctx.Response.Cookies.Append(DrawingQC.Web.Auth.CookieName, "",
        new CookieOptions { Expires = DateTimeOffset.UnixEpoch, Path = "/" });

// ---------- Auth request bodies ----------
record RegisterDto(string? username, string? email, string? password, string? name, string? role, string? securityQuestion, string? securityAnswer);
record ResetDto(string? email, string? answer, string? password);
record SecurityDto(string? question, string? answer);
record LoginDto(string? login, string? password, bool remember);
record ChangePwDto(string? current, string? next);
record ProfileDto(string? name, string? email, string? role, string? avatar);
record AdminRoleDto(string? id, string? role);
record AdminResetDto(string? id, string? password);
record AdminIdDto(string? id);
record RegToggleDto(bool open);