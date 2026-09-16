# Support Automation (DrawingQC)

Web app for the drawing-office tools: QC Check, KBR Tagwise Delivery Report, ConsList (S2NERGY), and the
Booklet generator. It runs on any machine with .NET, and is deployed to Render as a Docker container.

Two features need software that only exists on an engineer's Windows PC:

| Feature | Needs | Where it runs when the site is hosted |
|---|---|---|
| Sync to AutoCAD / Plant 3D | a running AutoCAD (COM) | the **Local Agent** on the user's PC |
| Booklet (Word template to PDF) | Microsoft Word (COM) | the **Local Agent** on the user's PC |

Everything else (accounts, QC Check, KBR report, ConsList) runs on the server.

---

## 1. Deploying to Render

### What you need

- The repo on GitHub (this one).
- A Render account with the GitHub repo connected.
- Nothing else: `render.yaml` at the repo root describes the whole setup.

### Steps

1. In Render, click **New > Blueprint**, pick this repository, branch `master`.
2. Render reads `render.yaml` and proposes two resources. Accept:
   - `drawingqc-db` — managed PostgreSQL (free plan).
   - `drawingqc-web` — Docker web service (free plan) built from `DrawingQC.Web/Dockerfile`.
3. Click **Apply**. The first build takes roughly 8 to 12 minutes because it also cross-compiles the
   Windows Local Agent and packs it into the image.
4. Open the service URL (`https://drawingqc-web.onrender.com` or whatever name Render assigned).
5. Register the first account. The first registered user becomes **Admin**. Then create the team's
   accounts from **User management**, or leave registration open for them.

Every later push to `master` redeploys automatically (`autoDeploy: true`).

### What `render.yaml` sets

| Key | Value | Why |
|---|---|---|
| `SUPPORTAUTOMATION_DB` | the database's internal connection string | Switches the app into PostgreSQL mode. Without it, the app stores everything in local JSON files, which do not survive a Render restart. |
| `healthCheckPath` | `/` | Render checks the home page. |
| `dockerContext` | `.` (repo root) | The web project links `DrawingQC.UI/QcEngine.cs` and the agent links files from `DrawingQC.Web`. |

### What is stored where

| Data | Storage on Render |
|---|---|
| Accounts, roles, registration setting | PostgreSQL |
| Session signing key | PostgreSQL (`settings` table), so users stay signed in across deploys |
| ConsList projects, file list, revisions | PostgreSQL |
| ConsList uploaded files and consolidated Excel/PDF | PostgreSQL (`blobs` table). The container disk is only a cache. |
| Per-tool run history | PostgreSQL |
| QC / KBR generated reports | in memory until the next restart; users download them right away |
| Booklets | on the user's PC (made by the Local Agent) |

### Free-plan limits to know

- **Free PostgreSQL is deleted 30 days after creation**, with a 14-day grace period. Upgrade
  `drawingqc-db` to the Basic plan before then if the data matters. The comment in `render.yaml` says the same.
- Free web services **spin down after 15 minutes without traffic** and take about a minute to wake.
  First request of the day is slow. This is normal.
- Free instances have **512 MB RAM and 0.1 CPU**. Very large QC zips or ConsList PDFs may be slow.
  The Starter plan removes most of this pain.
- Uploaded ConsList files count against the database's storage. Keep an eye on database size if
  the team uploads many large PDFs.

### Checking a deploy

1. Render **Logs** should show `[Db] PostgreSQL connected; schema ready.` on start.
2. `https://<your-site>/api/agent/info` should return JSON with `"downloadUrl": "/agent/DrawingQC.Agent.zip"`.
3. Sign in and open any tool. The sidebar shows a **Local agent** panel.

---

## 2. The Local Agent (for each engineer's PC)

The agent is a single `DrawingQC.Agent.exe` (about 50 MB, .NET bundled, nothing to install).

1. On the site, click **Download agent (Windows)** in the sidebar, or go to
   `https://<your-site>/agent/DrawingQC.Agent.zip`.
2. Unzip anywhere, for example `C:\DrawingQC-Agent`.
3. Double-click `DrawingQC.Agent.exe`. Keep the console window open while working.
   Windows SmartScreen may ask once: **More info > Run anyway**.
4. Reload the site. The sidebar panel turns green and lists the AutoCAD product it found and whether Word is installed.
5. Chrome or Edge may show a one-time **"allow this site to access your local network"** prompt. Click **Allow**.

Then **Sync to AutoCAD** and **Generate booklet** work exactly as in the desktop version. The booklet
is built on the PC and downloaded from there. Nothing leaves the PC except the run record
(file names, rev, date, size) that shows in the shared history.

Options:

```
DrawingQC.Agent.exe --port 5081
DrawingQC.Agent.exe --origin https://other-site.example.com     (repeatable)
```

By default the agent accepts requests only from `localhost` and `*.onrender.com`. If the site is
ever hosted somewhere else (for example the office VM behind Caddy), start the agent with
`--origin https://that-host` or set `DRAWINGQC_AGENT_ORIGINS`.

Full notes are in `DrawingQC.Agent/README.txt`, which is also inside the zip.

---

## 3. Running locally (Windows, for development)

```
dotnet run --project DrawingQC.Web
```

Opens on `http://localhost:5080`. Without `SUPPORTAUTOMATION_DB`, data is stored as JSON under
`%APPDATA%\SupportAutomation`. On Windows the server itself has AutoCAD and Word access, so the
agent is optional.

To test PostgreSQL mode locally:

```
docker run -d --name dqc-pg -e POSTGRES_PASSWORD=pg -e POSTGRES_DB=dqc -p 55432:5432 postgres:16
set SUPPORTAUTOMATION_DB=postgres://postgres:pg@localhost:55432/dqc
dotnet run --project DrawingQC.Web
```

On first start with an empty database the app imports any existing local JSON data.

To build the Render image locally:

```
docker build -f DrawingQC.Web/Dockerfile -t drawingqc .
docker run -p 8080:8080 -e PORT=8080 drawingqc
```

---

## 4. Alternative: hosting on the office VM behind Caddy

The same image (or `dotnet DrawingQC.Web.dll` directly) runs on the VM. Advantages over Render's free
plan: persistent disk, no spin-down, more memory, no database expiry.

Minimal Caddyfile:

```
support.example.com {
    reverse_proxy 127.0.0.1:8080
}
```

Set `PORT=8080` and `SUPPORTAUTOMATION_DB` for the app. Users then start the agent with
`--origin https://support.example.com`.

---

## 5. Project layout

| Path | What |
|---|---|
| `DrawingQC.Web/` | ASP.NET Core web app and UI (`wwwroot/index.html`) |
| `DrawingQC.Web/Dockerfile` | Render build: publishes the web app and the Windows agent |
| `DrawingQC.Agent/` | Local Agent (links `Booklet.cs`, `AutoCadSync.cs` from the web project) |
| `DrawingQC.UI/` | Original desktop app; `QcEngine.cs` is shared with the web app |
| `render.yaml` | Render Blueprint (web service + PostgreSQL) |
