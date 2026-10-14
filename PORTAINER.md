# Deploy on a Linux VM with Portainer

The whole app runs as a two-container stack: the C# web app + PostgreSQL. All data
(accounts, ConsList metadata, and the uploaded Excel/PDF files as blobs) lives in Postgres,
so the single `dqc_db` volume is the only thing to back up.

## Deploy (Portainer → Stacks → Add stack → Git repository)

1. **Repository URL:** `https://github.com/pushpalata-patil/DrawingQC`
2. **Compose path:** `docker-compose.yml`
3. **Environment variables** (Add environment variable — at minimum the password):
   ```
   POSTGRES_PASSWORD   = <a strong password>
   POSTGRES_USER       = sa_app            (optional, default)
   POSTGRES_DB         = supportautomation (optional, default)
   WEB_PORT            = 8080              (host port to publish; optional)
   ```
4. **Deploy the stack.** The first build takes ~8–12 min (it also cross-compiles the
   Windows Local Agent and packs it into the image). Portainer then starts `db`, waits for it
   to be healthy, and starts `web`.

> Portainer must have access to build images (it builds `web` from `DrawingQC.Web/Dockerfile`).
> Updating later: in the stack, **Pull and redeploy** to rebuild from the latest `master`.

## First launch

- Open `http://<vm-ip>:<WEB_PORT>` (default `:8080`).
- **Register the first account** → it becomes **Admin**. Create the team's accounts from
  User management.
- Backend log (`web` container) should show: `[Db] PostgreSQL connected; schema ready.`
- `http://<vm-ip>:<WEB_PORT>/api/agent/info` returns the agent download URL.

## What runs where

| Tool | Where |
|---|---|
| QC Check, Tagwise, ConsList, accounts | on the VM (these containers) |
| Booklet, AutoCAD sync | via the **Local Agent** (`DrawingQC.Agent.exe`) on a Windows PC — download it from the app sidebar |

Booklet/AutoCAD need Windows + Word/AutoCAD (COM), which a Linux VM can't provide — the agent
runs those on the user's own machine while everything else stays on the server.

## Production notes

- **Backups:** `docker exec <db-container> pg_dump -U sa_app supportautomation > backup.sql`,
  and back up the `dqc_db` volume. That captures everything, including ConsList files.
- **HTTPS / friendly URL:** put a reverse proxy (nginx / Caddy / Traefik / Cloudflare Tunnel)
  in front of `WEB_PORT`. The app serves plain HTTP inside the stack.
- **Resources:** a small instance is fine; give it enough RAM for document processing
  (≈1 GB+ recommended for large QC zips / ConsList PDFs). No spin-down, unlike free PaaS.
