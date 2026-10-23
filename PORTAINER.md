# Deploy on a Linux VM with Portainer

Pick the version that matches your database:

| Compose file | Database | Files | Use when |
|---|---|---|---|
| **`docker-compose.supabase.yml`** ⭐ | **Supabase** (external) | on the VM disk (`dqc_files`) | you use Supabase — web app only, no DB container to run |
| `docker-compose.yml` | PostgreSQL container on the VM | on the VM disk (`dqc_files`) | you want everything self-contained on the VM |

Either way, **ConsList files live on the VM disk** (`dqc_files` volume), so they never fill a small
managed DB.

---

## A) Supabase database (recommended) — `docker-compose.supabase.yml`

**Portainer → Stacks → Add stack → Git repository**

1. **Repository URL:** `https://github.com/pushpalata-patil/DrawingQC`
2. **Compose path:** `docker-compose.supabase.yml`
3. **Environment variables** (Add environment variable):
   ```
   SUPPORTAUTOMATION_DB = postgresql://postgres.<ref>:<password>@aws-0-<region>.pooler.supabase.com:5432/postgres
   WEB_PORT             = 8080   (optional; host port to publish)
   ```
   (Use the Supabase **Session pooler** string — see `SUPABASE.md`.)
4. **Deploy.** First build ~8–12 min. Only the `web` container runs; the DB is Supabase, files land
   on the `dqc_files` volume.

> ⚠️ Make the VM your **single** backend. Because files here live on disk (not in the DB), don't
> also run the Render web app against the same Supabase — stop/retire it, or the two instances
> disagree about where files are.

---

## B) Self-contained (Postgres on the VM) — `docker-compose.yml`

**Portainer → Stacks → Add stack → Git repository**

1. **Repository URL:** `https://github.com/pushpalata-patil/DrawingQC`
2. **Compose path:** `docker-compose.yml`
3. **Environment variables** (at minimum the password):
   ```
   POSTGRES_PASSWORD   = <a strong password>
   POSTGRES_USER       = sa_app            (optional, default)
   POSTGRES_DB         = supportautomation (optional, default)
   WEB_PORT            = 8080              (host port to publish; optional)
   ```
4. **Deploy the stack.** First build ~8–12 min. Portainer starts `db`, waits for it to be healthy,
   then starts `web`.

> Portainer must be able to build images (it builds `web` from `DrawingQC.Web/Dockerfile`).
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

- **Backups:**
  - *Supabase version:* the database is backed up by Supabase; back up the **`dqc_files`** volume
    (your uploaded files) e.g. `docker run --rm -v dqc_files:/d -v $PWD:/b alpine tar czf /b/files.tgz -C /d .`
  - *Self-contained version:* `docker exec <db-container> pg_dump -U sa_app supportautomation > backup.sql`
    **and** back up both the `dqc_db` and `dqc_files` volumes.
- **HTTPS / friendly URL:** put a reverse proxy (nginx / Caddy / Traefik / Cloudflare Tunnel)
  in front of `WEB_PORT`. The app serves plain HTTP inside the stack.
- **Resources:** a small instance is fine; give it enough RAM for document processing
  (≈1 GB+ recommended for large QC zips / ConsList PDFs). No spin-down, unlike free PaaS.
