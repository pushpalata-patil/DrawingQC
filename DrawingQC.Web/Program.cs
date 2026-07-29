using System.Net;
using System.Text;
using DrawingQC.Web;
using Microsoft.AspNetCore.Http.Features;

var builder = WebApplication.CreateBuilder(args);

// Allow large zip uploads (up to 500 MB).
const long MaxUpload = 500L * 1024 * 1024;
builder.Services.Configure<FormOptions>(o => o.MultipartBodyLengthLimit = MaxUpload);
builder.WebHost.ConfigureKestrel(o => o.Limits.MaxRequestBodySize = MaxUpload);

// Render provides the port to listen on via the PORT environment variable.
var port = Environment.GetEnvironmentVariable("PORT") ?? "8080";
builder.WebHost.UseUrls($"http://0.0.0.0:{port}");

var app = builder.Build();

app.MapGet("/", () => Results.Content(UploadPage(), "text/html"));
app.MapGet("/health", () => Results.Ok("ok"));

app.MapPost("/analyze", async (HttpRequest req) =>
{
    if (!req.HasFormContentType)
        return Results.BadRequest("Expected a file upload.");

    var form = await req.ReadFormAsync();
    var file = form.Files.GetFile("zip");
    if (file is null || file.Length == 0)
        return Results.Content(ErrorPage("Please choose a .zip file to upload."), "text/html");

    if (!file.FileName.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
        return Results.Content(ErrorPage("The uploaded file must be a .zip archive."), "text/html");

    try
    {
        // Buffer the upload so the zip is seekable.
        using var ms = new MemoryStream();
        await file.CopyToAsync(ms);
        ms.Position = 0;

        var results = QcEngine.Analyze(ms);
        var excel = QcEngine.BuildReportBytes(results);

        var wantsExcel = string.Equals(form["download"], "excel", StringComparison.OrdinalIgnoreCase);
        if (wantsExcel)
        {
            var outName = Path.GetFileNameWithoutExtension(file.FileName) + " - QC Report.xlsx";
            return Results.File(excel,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", outName);
        }

        return Results.Content(ResultsPage(file.FileName, results, excel), "text/html");
    }
    catch (Exception ex)
    {
        return Results.Content(ErrorPage("Could not process the zip: " + WebUtility.HtmlEncode(ex.Message)), "text/html");
    }
});

app.Run();

// ---------------- HTML rendering ----------------

static string UploadPage() => Shell("Drawing QC", $@"
  <h1>Drawing QC</h1>
  <p class='sub'>Upload a <b>.zip</b> of drawing PDFs. Each PDF's file name is checked against the
     drawing numbers actually printed inside the PDF. Duplicates across the set are flagged too.</p>
  <form method='post' action='/analyze' enctype='multipart/form-data' class='card'>
     <input type='file' name='zip' accept='.zip' required />
     <button type='submit'>Run QC</button>
  </form>
  {Legend()}");

static string ResultsPage(string fileName, List<PdfResult> results, byte[] excel)
{
    int matched = results.Count(r => r.FinalStatus == "Matched");
    int unmatched = results.Count(r => r.FinalStatus == "Unmatched");
    int duplicate = results.Count(r => r.FinalStatus == "Duplicate");
    string b64 = Convert.ToBase64String(excel);
    string outName = WebUtility.HtmlEncode(Path.GetFileNameWithoutExtension(fileName) + " - QC Report.xlsx");

    var sb = new StringBuilder();
    sb.Append($@"
      <h1>Drawing QC results</h1>
      <p class='sub'>File: <b>{WebUtility.HtmlEncode(fileName)}</b></p>
      <div class='counts'>
        <span class='pill total'>Total: {results.Count}</span>
        <span class='pill matched'>Matched: {matched}</span>
        <span class='pill unmatched'>Unmatched: {unmatched}</span>
        <span class='pill duplicate'>Duplicate: {duplicate}</span>
      </div>
      <p>
        <a class='btn' download='{outName}'
           href='data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}'>Download Excel</a>
        <a class='btn ghost' href='/'>New upload</a>
      </p>
      <table>
        <thead><tr>
          <th>Sr.No</th><th>Support PDF</th><th>1st drawing name</th>
          <th>2nd drawing name</th><th>Status</th><th>Not found in PDF</th>
        </tr></thead><tbody>");

    int sr = 1;
    foreach (var r in results)
    {
        string cls = r.FinalStatus.ToLowerInvariant();
        sb.Append($@"<tr class='{cls}'>
            <td>{sr++}</td>
            <td>{WebUtility.HtmlEncode(r.FileName)}</td>
            <td>{WebUtility.HtmlEncode(string.IsNullOrWhiteSpace(r.Drawing1) ? "-" : r.Drawing1)}</td>
            <td>{WebUtility.HtmlEncode(string.IsNullOrWhiteSpace(r.Drawing2) ? "-" : r.Drawing2)}</td>
            <td>{r.FinalStatus}</td>
            <td>{WebUtility.HtmlEncode(string.IsNullOrWhiteSpace(r.NotFoundInPdf) ? "-" : r.NotFoundInPdf)}</td>
          </tr>");
    }
    sb.Append("</tbody></table>");
    return Shell("Drawing QC results", sb.ToString());
}

static string ErrorPage(string message) => Shell("Drawing QC", $@"
  <h1>Drawing QC</h1>
  <div class='err'>{message}</div>
  <p><a class='btn' href='/'>Back</a></p>");

static string Legend() => @"
  <div class='legend'>
    <span class='pill matched'>Matched</span> both drawing names are found inside the PDF
    &nbsp;&nbsp;<span class='pill unmatched'>Unmatched</span> a drawing name is not found inside
    &nbsp;&nbsp;<span class='pill duplicate'>Duplicate</span> a drawing number appears more than once
  </div>";

static string Shell(string title, string body) => $@"<!doctype html>
<html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>
<title>{title}</title>
<style>
  :root {{ --dark:#37475a; --green:#c6efce; --greent:#006100; --red:#ffc7ce; --redt:#9c0006;
           --amber:#ffeb9c; --ambert:#9c6300; }}
  * {{ box-sizing:border-box; }}
  body {{ font-family:'Segoe UI',system-ui,Arial,sans-serif; margin:0; background:#f4f6f8; color:#222; }}
  .wrap {{ max-width:1100px; margin:0 auto; padding:24px; }}
  h1 {{ margin:0 0 6px; }}
  .sub {{ color:#555; margin:0 0 18px; }}
  .card {{ background:#fff; border:1px solid #e2e6ea; border-radius:10px; padding:22px; display:flex; gap:12px; align-items:center; }}
  input[type=file] {{ flex:1; }}
  button, .btn {{ background:var(--dark); color:#fff; border:0; border-radius:8px; padding:10px 18px;
           font-size:14px; cursor:pointer; text-decoration:none; display:inline-block; }}
  .btn.ghost {{ background:#fff; color:var(--dark); border:1px solid var(--dark); }}
  .counts {{ margin:12px 0; display:flex; gap:8px; flex-wrap:wrap; }}
  .pill {{ border-radius:999px; padding:4px 12px; font-size:13px; font-weight:600; }}
  .pill.total {{ background:#e7ebef; color:#37475a; }}
  .pill.matched {{ background:var(--green); color:var(--greent); }}
  .pill.unmatched {{ background:var(--red); color:var(--redt); }}
  .pill.duplicate {{ background:var(--amber); color:var(--ambert); }}
  .legend {{ margin-top:18px; color:#555; font-size:13px; }}
  .err {{ background:var(--red); color:var(--redt); padding:14px 16px; border-radius:8px; margin:12px 0; }}
  table {{ width:100%; border-collapse:collapse; margin-top:14px; background:#fff; font-size:13px; }}
  th, td {{ padding:7px 10px; border-bottom:1px solid #eef1f4; text-align:left; vertical-align:top; }}
  thead th {{ background:var(--dark); color:#fff; position:sticky; top:0; }}
  tr.matched td {{ background:var(--green); color:var(--greent); }}
  tr.unmatched td {{ background:var(--red); color:var(--redt); }}
  tr.duplicate td {{ background:var(--amber); color:var(--ambert); }}
</style></head>
<body><div class='wrap'>{body}</div></body></html>";
