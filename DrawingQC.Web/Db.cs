using System.Text.Json;
using Npgsql;

namespace DrawingQC.Web;

/// <summary>
/// Optional PostgreSQL backing store for the whole app. It is a drop-in, opt-in layer:
///   • No connection string configured  -> Db.Enabled is false and every feature keeps using the
///     existing file-based storage (users.json, ConsList manifests, etc.). Nothing changes locally.
///   • SUPPORTAUTOMATION_DB set (a shared server) -> the schema is created and, on the very first
///     run, the existing JSON data (accounts + ConsList metadata) is imported so nothing is lost.
///
/// Large Excel/PDF blobs stay on disk; the database holds the metadata and run history that a team
/// needs to share. The connection string is read from the SUPPORTAUTOMATION_DB environment variable,
/// e.g. "Host=localhost;Port=5432;Database=supportautomation;Username=sa_app;Password=…".
/// </summary>
public static class Db
{
    private static readonly string? ConnString = Normalize(Environment.GetEnvironmentVariable("SUPPORTAUTOMATION_DB"));

    /// <summary>True when a PostgreSQL connection string is configured; otherwise the app stays file-based.</summary>
    public static bool Enabled => !string.IsNullOrWhiteSpace(ConnString);

    // Accept both Npgsql key-value strings and the postgres://user:pass@host:port/db URL form that
    // managed hosts (Render, Railway, Heroku, …) hand out, so SUPPORTAUTOMATION_DB can be either.
    private static string? Normalize(string? cs)
    {
        if (string.IsNullOrWhiteSpace(cs)) return cs;
        if (!cs.StartsWith("postgres://", StringComparison.OrdinalIgnoreCase) &&
            !cs.StartsWith("postgresql://", StringComparison.OrdinalIgnoreCase))
            return cs;
        try
        {
            var uri = new Uri(cs);
            var userInfo = uri.UserInfo.Split(':', 2);
            var b = new NpgsqlConnectionStringBuilder
            {
                Host = uri.Host,
                Port = uri.Port > 0 ? uri.Port : 5432,
                Username = Uri.UnescapeDataString(userInfo[0]),
                Password = userInfo.Length > 1 ? Uri.UnescapeDataString(userInfo[1]) : "",
                Database = uri.AbsolutePath.TrimStart('/'),
                SslMode = SslMode.Prefer,
                TrustServerCertificate = true,
            };
            return b.ConnectionString;
        }
        catch { return cs; }
    }

    public static NpgsqlConnection Open()
    {
        var con = new NpgsqlConnection(ConnString);
        con.Open();
        return con;
    }

    private static string DataDir()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        if (string.IsNullOrWhiteSpace(appData)) appData = AppContext.BaseDirectory;
        return Path.Combine(appData, "SupportAutomation");
    }

    // Full whole-app schema: accounts, settings, ConsList metadata, and per-tool run history.
    private const string SchemaSql = @"
CREATE TABLE IF NOT EXISTS users (
    id                    TEXT PRIMARY KEY,
    username              TEXT NOT NULL,
    email                 TEXT NOT NULL,
    name                  TEXT NOT NULL DEFAULT '',
    role                  TEXT NOT NULL DEFAULT 'User',
    password_hash         TEXT NOT NULL,
    salt                  TEXT NOT NULL,
    avatar                TEXT,
    security_question     TEXT,
    security_answer_hash  TEXT,
    security_answer_salt  TEXT,
    created_at            TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_users_email    ON users (lower(email));
CREATE UNIQUE INDEX IF NOT EXISTS ux_users_username ON users (lower(username));

CREATE TABLE IF NOT EXISTS settings (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

-- File contents that must survive an ephemeral filesystem (Render, containers): ConsList source
-- uploads and the consolidated Excel/PDF per platform+category. key = path relative to the
-- ConsList root, forward slashes (e.g. RP5S/Internal/sources/<id>.pdf).
CREATE TABLE IF NOT EXISTS blobs (
    key        TEXT PRIMARY KEY,
    data       BYTEA NOT NULL,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS conslist_projects (
    name TEXT PRIMARY KEY,
    ord  INT NOT NULL DEFAULT 0
);

-- Per platform+category state. category is Internal | External | Combined.
CREATE TABLE IF NOT EXISTS conslist_category (
    platform        TEXT NOT NULL,
    category        TEXT NOT NULL,
    excel_rev       INT  NOT NULL DEFAULT 0,
    pdf_rev         INT  NOT NULL DEFAULT 0,
    last_excel_date TEXT NOT NULL DEFAULT '',
    PRIMARY KEY (platform, category)
);

-- One row per uploaded file (its bytes live on disk under sources\<id><ext>).
CREATE TABLE IF NOT EXISTS conslist_files (
    id         TEXT PRIMARY KEY,
    platform   TEXT NOT NULL,
    category   TEXT NOT NULL,   -- Internal | External
    kind       TEXT NOT NULL,   -- excel | pdf
    ext        TEXT NOT NULL,
    file_date  TEXT NOT NULL,   -- dd-MM-yyyy
    name       TEXT NOT NULL,
    cnt        INT  NOT NULL DEFAULT 0,
    ord        INT  NOT NULL DEFAULT 0,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS ix_conslist_files_pc ON conslist_files (platform, category, kind, ord);

-- Per-tool run history (who ran what, when, and the result summary).
CREATE TABLE IF NOT EXISTS qc_runs (
    id          TEXT PRIMARY KEY,
    user_id     TEXT,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    source_name TEXT,
    total       INT, matched INT, unmatched INT, duplicate INT,
    report_name TEXT
);

CREATE TABLE IF NOT EXISTS tagreport_runs (
    id          TEXT PRIMARY KEY,
    user_id     TEXT,
    created_at  TIMESTAMPTZ NOT NULL DEFAULT now(),
    total       INT, delivered INT, pending INT, done_files INT,
    sheet       TEXT, column_name TEXT, report_name TEXT
);

CREATE TABLE IF NOT EXISTS booklet_runs (
    id            TEXT PRIMARY KEY,
    user_id       TEXT,
    created_at    TIMESTAMPTZ NOT NULL DEFAULT now(),
    template_name TEXT, excel_name TEXT, drawings_name TEXT, bom_name TEXT,
    output_name   TEXT, rev TEXT, doc_date TEXT, size_mb DOUBLE PRECISION
);

CREATE TABLE IF NOT EXISTS mto_runs (
    id         TEXT PRIMARY KEY,
    user_id    TEXT,
    created_at TIMESTAMPTZ NOT NULL DEFAULT now(),
    summary    TEXT
);
";

    /// <summary>Create the schema and import existing JSON data on first run. No-op when not configured.</summary>
    public static void Init()
    {
        if (!Enabled)
        {
            Console.WriteLine("[Db] SUPPORTAUTOMATION_DB not set — using local file storage.");
            return;
        }
        using var con = Open();
        using (var cmd = new NpgsqlCommand(SchemaSql, con)) cmd.ExecuteNonQuery();
        Console.WriteLine("[Db] PostgreSQL connected; schema ready.");
        ImportFromJsonIfEmpty(con);
    }

    // timestamptz columns require a UTC DateTime under Npgsql; normalise whatever the JSON gave us.
    private static DateTime Utc(DateTime d) => d.Kind switch
    {
        DateTimeKind.Utc => d,
        DateTimeKind.Local => d.ToUniversalTime(),
        _ => DateTime.SpecifyKind(d, DateTimeKind.Utc),
    };

    private static long Count(NpgsqlConnection con, string table)
    {
        using var cmd = new NpgsqlCommand($"SELECT count(*) FROM {table}", con);
        return Convert.ToInt64(cmd.ExecuteScalar() ?? 0L);
    }

    private static void Exec(NpgsqlConnection con, string sql, params (string name, object value)[] ps) => Exec(con, null, sql, ps);

    private static void Exec(NpgsqlConnection con, NpgsqlTransaction? tx, string sql, params (string name, object value)[] ps)
    {
        using var cmd = new NpgsqlCommand(sql, con, tx);
        foreach (var (n, v) in ps) cmd.Parameters.AddWithValue(n, v ?? DBNull.Value);
        cmd.ExecuteNonQuery();
    }

    // ---------- accounts (used by Auth when Db.Enabled) ----------

    private const string UserCols =
        "id,username,email,name,role,password_hash,salt,avatar,security_question,security_answer_hash,security_answer_salt,created_at";

    public static List<UserAccount> LoadUsers()
    {
        var list = new List<UserAccount>();
        using var con = Open();
        using var cmd = new NpgsqlCommand($"SELECT {UserCols} FROM users ORDER BY created_at", con);
        using var r = cmd.ExecuteReader();
        while (r.Read())
            list.Add(new UserAccount
            {
                Id = r.GetString(0),
                Username = r.GetString(1),
                Email = r.GetString(2),
                Name = r.GetString(3),
                Role = r.GetString(4),
                PasswordHash = r.GetString(5),
                Salt = r.GetString(6),
                Avatar = r.IsDBNull(7) ? null : r.GetString(7),
                SecurityQuestion = r.IsDBNull(8) ? null : r.GetString(8),
                SecurityAnswerHash = r.IsDBNull(9) ? null : r.GetString(9),
                SecurityAnswerSalt = r.IsDBNull(10) ? null : r.GetString(10),
                CreatedAt = r.GetDateTime(11),
            });
        return list;
    }

    // Make the users table exactly match the given list (upsert present rows, delete the rest) —
    // mirrors the file store's whole-list Save so Auth's existing logic is unchanged.
    public static void SaveUsers(List<UserAccount> users)
    {
        using var con = Open();
        using var tx = con.BeginTransaction();
        var ids = users.Select(u => u.Id).ToArray();
        Exec(con, tx, "DELETE FROM users WHERE NOT (id = ANY(@ids))", ("ids", ids));
        foreach (var u in users)
            Exec(con, tx, @"INSERT INTO users(id,username,email,name,role,password_hash,salt,avatar,security_question,security_answer_hash,security_answer_salt,created_at)
                            VALUES(@id,@un,@em,@nm,@ro,@ph,@sa,@av,@sq,@sah,@sas,@ca)
                            ON CONFLICT (id) DO UPDATE SET
                              username=excluded.username, email=excluded.email, name=excluded.name, role=excluded.role,
                              password_hash=excluded.password_hash, salt=excluded.salt, avatar=excluded.avatar,
                              security_question=excluded.security_question, security_answer_hash=excluded.security_answer_hash,
                              security_answer_salt=excluded.security_answer_salt",
                ("id", u.Id), ("un", u.Username), ("em", u.Email), ("nm", u.Name), ("ro", u.Role),
                ("ph", u.PasswordHash), ("sa", u.Salt), ("av", (object?)u.Avatar ?? DBNull.Value),
                ("sq", (object?)u.SecurityQuestion ?? DBNull.Value), ("sah", (object?)u.SecurityAnswerHash ?? DBNull.Value),
                ("sas", (object?)u.SecurityAnswerSalt ?? DBNull.Value), ("ca", Utc(u.CreatedAt)));
        tx.Commit();
    }

    // ---------- key/value settings ----------

    public static string? GetSetting(string key)
    {
        using var con = Open();
        using var cmd = new NpgsqlCommand("SELECT value FROM settings WHERE key=@k", con);
        cmd.Parameters.AddWithValue("k", key);
        return cmd.ExecuteScalar() as string;
    }

    public static void SetSetting(string key, string value)
    {
        using var con = Open();
        Exec(con, "INSERT INTO settings(key,value) VALUES(@k,@v) ON CONFLICT (key) DO UPDATE SET value=excluded.value", ("k", key), ("v", value));
    }

    // ---------- blobs (file bytes that must outlive the container's disk) ----------

    public static byte[]? GetBlob(string key)
    {
        using var con = Open();
        using var cmd = new NpgsqlCommand("SELECT data FROM blobs WHERE key=@k", con);
        cmd.Parameters.AddWithValue("k", key);
        return cmd.ExecuteScalar() as byte[];
    }

    public static bool HasBlob(string key)
    {
        using var con = Open();
        using var cmd = new NpgsqlCommand("SELECT 1 FROM blobs WHERE key=@k", con);
        cmd.Parameters.AddWithValue("k", key);
        return cmd.ExecuteScalar() != null;
    }

    public static void PutBlob(string key, byte[] data)
    {
        using var con = Open();
        Exec(con, @"INSERT INTO blobs(key,data,updated_at) VALUES(@k,@d,now())
                    ON CONFLICT (key) DO UPDATE SET data=excluded.data, updated_at=now()", ("k", key), ("d", data));
    }

    public static void DeleteBlob(string key)
    {
        using var con = Open();
        Exec(con, "DELETE FROM blobs WHERE key=@k", ("k", key));
    }

    public static void DeleteBlobsWithPrefix(string prefix)
    {
        using var con = Open();
        Exec(con, "DELETE FROM blobs WHERE key LIKE @p", ("p", prefix.Replace("%", "\\%").Replace("_", "\\_") + "%"));
    }

    // One-time migration of the existing on-disk JSON into the database (only when the DB is empty).
    private static void ImportFromJsonIfEmpty(NpgsqlConnection con)
    {
        if (Count(con, "users") > 0 || Count(con, "conslist_projects") > 0) return; // already populated
        var dir = DataDir();
        int users = 0, files = 0, projects = 0;

        // Accounts
        var usersFile = Path.Combine(dir, "users.json");
        if (File.Exists(usersFile))
        {
            List<UserAccount> list;
            try { list = JsonSerializer.Deserialize<List<UserAccount>>(File.ReadAllText(usersFile)) ?? new(); }
            catch { list = new(); }
            foreach (var u in list)
            {
                Exec(con, @"INSERT INTO users(id,username,email,name,role,password_hash,salt,avatar,security_question,security_answer_hash,security_answer_salt,created_at)
                            VALUES(@id,@un,@em,@nm,@ro,@ph,@sa,@av,@sq,@sah,@sas,@ca) ON CONFLICT (id) DO NOTHING",
                    ("id", u.Id), ("un", u.Username), ("em", u.Email), ("nm", u.Name), ("ro", u.Role),
                    ("ph", u.PasswordHash), ("sa", u.Salt), ("av", (object?)u.Avatar ?? DBNull.Value),
                    ("sq", (object?)u.SecurityQuestion ?? DBNull.Value), ("sah", (object?)u.SecurityAnswerHash ?? DBNull.Value),
                    ("sas", (object?)u.SecurityAnswerSalt ?? DBNull.Value), ("ca", Utc(u.CreatedAt)));
                users++;
            }
        }

        // Settings (registration toggle)
        var settingsFile = Path.Combine(dir, "settings.json");
        if (File.Exists(settingsFile))
        {
            try
            {
                using var d = JsonDocument.Parse(File.ReadAllText(settingsFile));
                if (d.RootElement.TryGetProperty("RegistrationOpen", out var ro))
                    Exec(con, "INSERT INTO settings(key,value) VALUES('RegistrationOpen',@v) ON CONFLICT (key) DO UPDATE SET value=excluded.value",
                        ("v", ro.GetBoolean() ? "true" : "false"));
            }
            catch { }
        }

        // ConsList projects + per-file metadata
        var cl = Path.Combine(dir, "ConsList");
        var projFile = Path.Combine(cl, "projects.json");
        if (File.Exists(projFile))
        {
            List<string> projs;
            try { projs = JsonSerializer.Deserialize<List<string>>(File.ReadAllText(projFile)) ?? new(); }
            catch { projs = new(); }
            for (int i = 0; i < projs.Count; i++)
            {
                Exec(con, "INSERT INTO conslist_projects(name,ord) VALUES(@n,@o) ON CONFLICT (name) DO NOTHING", ("n", projs[i]), ("o", i));
                projects++;
            }
            foreach (var plat in projs)
            {
                foreach (var cat in new[] { "Internal", "External" })
                {
                    var mf = Path.Combine(cl, plat, cat, "manifest.json");
                    if (!File.Exists(mf)) continue;
                    CatManifest m;
                    try { m = JsonSerializer.Deserialize<CatManifest>(File.ReadAllText(mf)) ?? new(); }
                    catch { continue; }
                    Exec(con, @"INSERT INTO conslist_category(platform,category,excel_rev,pdf_rev,last_excel_date)
                                VALUES(@p,@c,@er,@pr,@d) ON CONFLICT (platform,category) DO NOTHING",
                        ("p", plat), ("c", cat), ("er", m.ExcelRev), ("pr", m.PdfRev), ("d", m.LastExcelDate ?? ""));
                    int ord = 0;
                    foreach (var e in m.ExcelEntries) { InsertFile(con, null, plat, cat, "excel", e, ord++); files++; }
                    ord = 0;
                    foreach (var e in m.PdfEntries) { InsertFile(con, null, plat, cat, "pdf", e, ord++); files++; }
                }
                // Combined-download revisions
                var cj = Path.Combine(cl, plat, "_combined.json");
                if (File.Exists(cj))
                {
                    try
                    {
                        using var d = JsonDocument.Parse(File.ReadAllText(cj));
                        int er = d.RootElement.TryGetProperty("ExcelRev", out var e1) ? e1.GetInt32() : 0;
                        int pr = d.RootElement.TryGetProperty("PdfRev", out var p1) ? p1.GetInt32() : 0;
                        Exec(con, @"INSERT INTO conslist_category(platform,category,excel_rev,pdf_rev,last_excel_date)
                                    VALUES(@p,'Combined',@er,@pr,'') ON CONFLICT (platform,category) DO NOTHING",
                            ("p", plat), ("er", er), ("pr", pr));
                    }
                    catch { }
                }
            }
        }
        Console.WriteLine($"[Db] Imported from JSON: {users} user(s), {projects} project(s), {files} ConsList file(s).");
    }

    private static void InsertFile(NpgsqlConnection con, NpgsqlTransaction? tx, string plat, string cat, string kind, ConsEntry e, int ord)
    {
        if (string.IsNullOrEmpty(e.Id)) e.Id = Guid.NewGuid().ToString("N");
        if (string.IsNullOrEmpty(e.Ext)) e.Ext = kind == "pdf" ? ".pdf" : ".xlsx";
        Exec(con, tx, @"INSERT INTO conslist_files(id,platform,category,kind,ext,file_date,name,cnt,ord)
                        VALUES(@id,@p,@c,@k,@x,@d,@n,@ct,@o) ON CONFLICT (id) DO NOTHING",
            ("id", e.Id), ("p", plat), ("c", cat), ("k", kind), ("x", e.Ext),
            ("d", e.Date), ("n", e.Name), ("ct", e.Count), ("o", ord));
    }

    // ---------- ConsList metadata (used by ConsList when Db.Enabled) ----------

    public static List<string> LoadProjects()
    {
        var list = new List<string>();
        using var con = Open();
        using var cmd = new NpgsqlCommand("SELECT name FROM conslist_projects ORDER BY ord", con);
        using var r = cmd.ExecuteReader();
        while (r.Read()) list.Add(r.GetString(0));
        return list;
    }

    public static void SaveProjects(List<string> names)
    {
        using var con = Open();
        using var tx = con.BeginTransaction();
        Exec(con, tx, "DELETE FROM conslist_projects");
        for (int i = 0; i < names.Count; i++)
            Exec(con, tx, "INSERT INTO conslist_projects(name,ord) VALUES(@n,@o)", ("n", names[i]), ("o", i));
        tx.Commit();
    }

    public static void DeletePlatformData(string platform)
    {
        using var con = Open();
        Exec(con, "DELETE FROM conslist_files WHERE platform=@p", ("p", platform));
        Exec(con, "DELETE FROM conslist_category WHERE platform=@p", ("p", platform));
    }

    public static CatManifest LoadManifest(string platform, string category)
    {
        var m = new CatManifest();
        using var con = Open();
        using (var cmd = new NpgsqlCommand("SELECT excel_rev,pdf_rev,last_excel_date FROM conslist_category WHERE platform=@p AND category=@c", con))
        {
            cmd.Parameters.AddWithValue("p", platform); cmd.Parameters.AddWithValue("c", category);
            using var r = cmd.ExecuteReader();
            if (r.Read()) { m.ExcelRev = r.GetInt32(0); m.PdfRev = r.GetInt32(1); m.LastExcelDate = r.IsDBNull(2) ? "" : r.GetString(2); }
        }
        using (var cmd = new NpgsqlCommand("SELECT id,ext,file_date,name,cnt,kind FROM conslist_files WHERE platform=@p AND category=@c ORDER BY ord", con))
        {
            cmd.Parameters.AddWithValue("p", platform); cmd.Parameters.AddWithValue("c", category);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                var e = new ConsEntry { Id = r.GetString(0), Ext = r.GetString(1), Date = r.GetString(2), Name = r.GetString(3), Count = r.GetInt32(4) };
                if (r.GetString(5) == "excel") m.ExcelEntries.Add(e); else m.PdfEntries.Add(e);
            }
        }
        return m;
    }

    public static void SaveManifest(string platform, string category, CatManifest m)
    {
        using var con = Open();
        using var tx = con.BeginTransaction();
        Exec(con, tx, @"INSERT INTO conslist_category(platform,category,excel_rev,pdf_rev,last_excel_date)
                        VALUES(@p,@c,@er,@pr,@d) ON CONFLICT (platform,category)
                        DO UPDATE SET excel_rev=excluded.excel_rev, pdf_rev=excluded.pdf_rev, last_excel_date=excluded.last_excel_date",
            ("p", platform), ("c", category), ("er", m.ExcelRev), ("pr", m.PdfRev), ("d", m.LastExcelDate ?? ""));
        Exec(con, tx, "DELETE FROM conslist_files WHERE platform=@p AND category=@c", ("p", platform), ("c", category));
        int ord = 0;
        foreach (var e in m.ExcelEntries) InsertFile(con, tx, platform, category, "excel", e, ord++);
        ord = 0;
        foreach (var e in m.PdfEntries) InsertFile(con, tx, platform, category, "pdf", e, ord++);
        tx.Commit();
    }

    public static (int excel, int pdf) LoadCombinedRev(string platform)
    {
        using var con = Open();
        using var cmd = new NpgsqlCommand("SELECT excel_rev,pdf_rev FROM conslist_category WHERE platform=@p AND category='Combined'", con);
        cmd.Parameters.AddWithValue("p", platform);
        using var r = cmd.ExecuteReader();
        return r.Read() ? (r.GetInt32(0), r.GetInt32(1)) : (0, 0);
    }

    public static void SaveCombinedRev(string platform, int excel, int pdf)
    {
        using var con = Open();
        Exec(con, @"INSERT INTO conslist_category(platform,category,excel_rev,pdf_rev,last_excel_date)
                    VALUES(@p,'Combined',@e,@f,'') ON CONFLICT (platform,category)
                    DO UPDATE SET excel_rev=excluded.excel_rev, pdf_rev=excluded.pdf_rev",
            ("p", platform), ("e", excel), ("f", pdf));
    }

    // ---------- per-tool run history (best-effort: never breaks a tool run) ----------

    private static void Record(string sql, params (string name, object value)[] ps)
    {
        if (!Enabled) return;
        try { using var con = Open(); Exec(con, sql, ps); }
        catch (Exception ex) { Console.Error.WriteLine("[Db] run-history insert failed: " + ex.Message); }
    }

    public static void RecordQcRun(string? userId, string? sourceName, int total, int matched, int unmatched, int duplicate, string? reportName) =>
        Record(@"INSERT INTO qc_runs(id,user_id,source_name,total,matched,unmatched,duplicate,report_name)
                 VALUES(@id,@u,@s,@t,@m,@un,@d,@r)",
            ("id", Guid.NewGuid().ToString("N")), ("u", (object?)userId ?? DBNull.Value), ("s", (object?)sourceName ?? DBNull.Value),
            ("t", total), ("m", matched), ("un", unmatched), ("d", duplicate), ("r", (object?)reportName ?? DBNull.Value));

    public static void RecordTagreportRun(string? userId, int total, int delivered, int pending, int doneFiles, string? sheet, string? column, string? reportName) =>
        Record(@"INSERT INTO tagreport_runs(id,user_id,total,delivered,pending,done_files,sheet,column_name,report_name)
                 VALUES(@id,@u,@t,@dl,@pn,@df,@sh,@col,@r)",
            ("id", Guid.NewGuid().ToString("N")), ("u", (object?)userId ?? DBNull.Value), ("t", total), ("dl", delivered),
            ("pn", pending), ("df", doneFiles), ("sh", (object?)sheet ?? DBNull.Value), ("col", (object?)column ?? DBNull.Value),
            ("r", (object?)reportName ?? DBNull.Value));

    public static void RecordBookletRun(string? userId, string? templateName, string? excelName, string? drawingsName, string? bomName, string? outputName, string? rev, string? docDate, double sizeMb) =>
        Record(@"INSERT INTO booklet_runs(id,user_id,template_name,excel_name,drawings_name,bom_name,output_name,rev,doc_date,size_mb)
                 VALUES(@id,@u,@tn,@en,@dn,@bn,@on,@rv,@dt,@sz)",
            ("id", Guid.NewGuid().ToString("N")), ("u", (object?)userId ?? DBNull.Value), ("tn", (object?)templateName ?? DBNull.Value),
            ("en", (object?)excelName ?? DBNull.Value), ("dn", (object?)drawingsName ?? DBNull.Value), ("bn", (object?)bomName ?? DBNull.Value),
            ("on", (object?)outputName ?? DBNull.Value), ("rv", (object?)rev ?? DBNull.Value), ("dt", (object?)docDate ?? DBNull.Value), ("sz", sizeMb));
}
