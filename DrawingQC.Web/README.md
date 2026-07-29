# DrawingQC.Web

Web version of the Drawing QC tool. Upload a `.zip` of drawing PDFs in the browser; it
checks each PDF's file name against the drawing numbers printed inside the PDF, flags
duplicates across the set, shows a colour-coded results table, and lets you download the
Excel report. Cross-platform (runs on Linux) so it can be hosted on Render.

## Run locally
```
cd DrawingQC.Web
dotnet run
# open the printed http://localhost:xxxx URL
```

## Deploy to Render (via GitHub + Docker)

1. Push this repository to a GitHub repo.
2. In Render: **New +** -> **Blueprint**, and select the repo. Render reads `render.yaml`
   at the repo root and creates a Docker web service `drawingqc-web` (free plan) that builds
   `DrawingQC.Web/Dockerfile`.
   - Alternatively: **New +** -> **Web Service** -> select the repo ->
     Runtime **Docker**, Dockerfile path `DrawingQC.Web/Dockerfile`,
     Docker build context `DrawingQC.Web`.
3. Deploy. When it's live, open the Render URL and upload a zip.

Notes:
- The app listens on the port from the `PORT` env var (Render sets this automatically).
- Health check path: `/health`.
- Max upload size: 500 MB.
- The free plan has 512 MB RAM and sleeps when idle (first request after idle is slow).
