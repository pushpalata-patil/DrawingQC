# Using Supabase (free Postgres) + migrating your data

The app talks to Postgres directly (Npgsql). Moving to Supabase needs **one credential — the
Postgres connection string.** You do **not** need the Supabase API keys (anon / service_role);
those are for Supabase's REST/JS client, which this app doesn't use.

> ## ⚠️ Important: keep files OUT of the database on the free tier
> Supabase free gives **~0.5 GB** of database. Your ConsList files are large (local data measured
> **~1.8 GB**). Storing those PDFs/Excels *in the DB* would blow past the free limit several times
> over. So the design here is: **files live on the server's disk (a Docker volume); only accounts +
> ConsList metadata + run history go in Supabase** (tiny). This is controlled by
> `SUPPORTAUTOMATION_FILE_STORE=disk` — already set in `docker-compose.yml`.

---

## 1. The only credential you need, and where to find it

1. Create a project at **https://supabase.com** (choose a region near your users; set a strong
   **database password** — save it).
2. In the project: **Settings** (gear, bottom-left) → **Database** → **Connection string**.
3. Choose the **Session pooler** tab ("Connection pooling", Mode = **Session**, **port 5432**).
   > ⚠️ Use the **pooler**, not "Direct connection". The direct host (`db.<ref>.supabase.co`) is
   > **IPv6-only** and fails on most servers. The session pooler is **IPv4** and persistent — correct here.
4. Copy it (replace `<YOUR-DB-PASSWORD>`):
   ```
   postgresql://postgres.<PROJECT_REF>:<YOUR-DB-PASSWORD>@aws-0-<region>.pooler.supabase.com:5432/postgres
   ```

That string is your **`SUPPORTAUTOMATION_DB`**. If the password has special characters (`@ : / ?`),
use the key-value form so they don't break the URL:
```
Host=aws-0-<region>.pooler.supabase.com;Port=5432;Database=postgres;Username=postgres.<PROJECT_REF>;Password=<YOUR-DB-PASSWORD>
```
Either form works — the app adds SSL automatically.

---

## 2. Env vars the app reads (no code changes needed)

| Variable | Value | Why |
|---|---|---|
| `SUPPORTAUTOMATION_DB` | the Supabase string above | switches on PostgreSQL mode |
| `SUPPORTAUTOMATION_FILE_STORE` | `disk` | keep ConsList files on the volume, not in the DB |
| `SUPPORTAUTOMATION_DATA_DIR` | `/data` | fixed path for the files volume (mounted in compose) |

For Portainer: in `docker-compose.yml`, delete the `db` service + `dqc_db` volume, drop `depends_on`,
and set `SUPPORTAUTOMATION_DB` to the Supabase string. Keep `SUPPORTAUTOMATION_FILE_STORE=disk`,
`SUPPORTAUTOMATION_DATA_DIR=/data`, and the `dqc_files` volume.

On first start the app **creates its own schema** and logs `[Db] PostgreSQL connected; schema ready.`

---

## 3. Migrate your current data — local first, then Supabase

Two separate things move: **metadata → the database**, and **files → the disk/volume**.

### Step 1 — metadata into a LOCAL Postgres first (verify it works)
```powershell
# create an empty local DB once:  psql -U postgres -c "CREATE DATABASE supportautomation;"
$env:SUPPORTAUTOMATION_DB = "Host=localhost;Port=5432;Database=supportautomation;Username=postgres;Password=<local-pw>"
$env:SUPPORTAUTOMATION_FILE_STORE = "disk"   # files stay in %APPDATA%, DB gets metadata only
cd C:\Users\pushpat\Downloads\dqc\DrawingQC.Web
dotnet run -c Release
# log:  [Db] Imported from JSON: N user(s), N project(s), N file(s), 0 file blob(s).
# open http://localhost:5080 -> log in -> ConsList downloads work (files read from %APPDATA%).
```

### Step 2 — metadata into SUPABASE
```powershell
$env:SUPPORTAUTOMATION_DB = "<supabase session-pooler string>"
$env:SUPPORTAUTOMATION_FILE_STORE = "disk"
dotnet run -c Release
# same import line -> Supabase now has accounts + ConsList metadata + run history (small).
```
> Import only runs on an **empty** target DB (no users, no ConsList projects).

### Step 3 — files onto the server volume
The files (`~1.8 GB`) must sit on the VM so the app can read them. Copy your local ConsList folder
into the `dqc_files` volume (which the app sees at `/data/ConsList`):
```bash
# on the VM, with the stack deployed (web container running):
#   source on your PC:  C:\Users\pushpat\AppData\Roaming\SupportAutomation\ConsList\
#   1) copy that folder to the VM (scp/rsync/WinSCP) -> e.g. /home/you/ConsList
#   2) copy it into the running web container's /data:
docker cp /home/you/ConsList <web-container>:/data/
#   (result: /data/ConsList/RP5S/... inside the container, on the dqc_files volume)
```
Restart the `web` container. ConsList now finds every file on the volume; metadata comes from Supabase.

---

## 4. Verify
1. App log: `PostgreSQL connected; schema ready.`
2. Log in with a migrated account (or register the first → Admin on a fresh DB).
3. ConsList → a project → **Download** the consolidated Excel → opens **with logos** (proves the
   files are present on the volume and metadata matches).
4. Supabase dashboard → **Database → Reports** → size should be small (MBs, not GBs).

## Notes
- **If you'd rather keep everything in one place (files in the DB too):** set
  `SUPPORTAUTOMATION_FILE_STORE=db` and the app stores files as blobs — but then you need a DB big
  enough (Supabase **Pro**, 8 GB, ~$25/mo), since ~1.8 GB won't fit the free tier.
- Supabase **pauses** a free project after ~1 week idle (resume from the dashboard).
- Keep the connection string in env vars — never commit it.
