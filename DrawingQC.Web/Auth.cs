using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace DrawingQC.Web;

/// <summary>A registered user account (stored on disk in users.json).</summary>
public sealed class UserAccount
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public string Username { get; set; } = "";
    public string Email { get; set; } = "";
    public string Name { get; set; } = "";
    public string Role { get; set; } = "User";
    public string PasswordHash { get; set; } = "";
    public string Salt { get; set; } = "";
    public string? Avatar { get; set; }            // small image as a data URL
    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;
}

/// <summary>
/// File-backed accounts + PBKDF2 password hashing + HMAC-signed session tokens.
/// Deliberately lightweight (this is a local single-instance tool), but passwords are
/// salted+hashed and never stored or returned in the clear, and session cookies are signed
/// so they survive app restarts without a server-side session table.
/// </summary>
public static class Auth
{
    public const string CookieName = "sa_session";

    // Store accounts in a stable per-user location (e.g. %APPDATA%\SupportAutomation on
    // Windows) so they survive rebuilds / clean / republish — NOT inside bin/.
    private static readonly string DataDir = ResolveDataDir();
    private static readonly string UsersFile = Path.Combine(DataDir, "users.json");
    private static readonly string SettingsFile = Path.Combine(DataDir, "settings.json");
    private static readonly string KeyFile = Path.Combine(DataDir, "auth.key");
    private static readonly object Gate = new();
    private static readonly byte[] Key = LoadOrCreateKey();
    private static readonly JsonSerializerOptions JsonOpts = new() { WriteIndented = true };

    private static string ResolveDataDir()
    {
        var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        if (string.IsNullOrWhiteSpace(appData)) appData = AppContext.BaseDirectory; // fallback
        var dir = Path.Combine(appData, "SupportAutomation");
        try { Directory.CreateDirectory(dir); } catch { }
        return dir;
    }

    // ---------- user store ----------

    private static List<UserAccount> Load()
    {
        lock (Gate)
        {
            if (!File.Exists(UsersFile)) return new();
            try { return JsonSerializer.Deserialize<List<UserAccount>>(File.ReadAllText(UsersFile)) ?? new(); }
            catch { return new(); }
        }
    }

    private static void Save(List<UserAccount> users)
    {
        lock (Gate)
        {
            Directory.CreateDirectory(DataDir);
            File.WriteAllText(UsersFile, JsonSerializer.Serialize(users, JsonOpts));
        }
    }

    public static UserAccount? FindById(string id) => Load().FirstOrDefault(u => u.Id == id);

    public static UserAccount? FindByLogin(string login)
    {
        login = (login ?? "").Trim();
        return Load().FirstOrDefault(u =>
            u.Username.Equals(login, StringComparison.OrdinalIgnoreCase) ||
            (!string.IsNullOrEmpty(u.Email) && u.Email.Equals(login, StringComparison.OrdinalIgnoreCase)));
    }

    // ---------- registration / login ----------

    public static (bool ok, string? error, UserAccount? user) Register(string? username, string? email, string? password, string? name, string? role)
    {
        username = (username ?? "").Trim();
        email = (email ?? "").Trim();

        var users = Load();
        bool first = users.Count == 0;
        // The very first account bootstraps the Admin. After that, sign-ups can be closed by
        // an admin, and no one can self-assign the Admin role.
        if (!first && !GetSettings().RegistrationOpen)
            return (false, "New sign-ups are currently disabled — please contact your administrator.", null);

        // Email is the login identifier. Username is optional/internal — derived from the
        // email when not supplied — so the sign-up form only needs full name + email.
        if (email.Length == 0 || !email.Contains('@')) return (false, "Please enter a valid email address.", null);
        if (string.IsNullOrEmpty(password) || password!.Length < 6) return (false, "Password must be at least 6 characters.", null);
        if (username.Length == 0) username = email;

        if (users.Any(u => u.Email.Equals(email, StringComparison.OrdinalIgnoreCase)))
            return (false, "That email is already registered.", null);
        if (users.Any(u => u.Username.Equals(username, StringComparison.OrdinalIgnoreCase)))
            return (false, "That username is already taken.", null);

        var finalRole = string.IsNullOrWhiteSpace(role) ? "User" : role!.Trim();
        if (first) finalRole = "Admin";                                                   // bootstrap
        else if (finalRole.Equals("Admin", StringComparison.OrdinalIgnoreCase)) finalRole = "User"; // no self-admin

        var salt = RandomNumberGenerator.GetBytes(16);
        var user = new UserAccount
        {
            Username = username,
            Email = email,
            Name = string.IsNullOrWhiteSpace(name) ? username : name!.Trim(),
            Role = finalRole,
            Salt = Convert.ToBase64String(salt),
            PasswordHash = Convert.ToBase64String(HashPassword(password!, salt)),
        };
        users.Add(user);
        Save(users);
        return (true, null, user);
    }

    public static UserAccount? Validate(string? login, string? password)
    {
        var user = FindByLogin(login ?? "");
        if (user == null) return null;
        var salt = Convert.FromBase64String(user.Salt);
        var hash = HashPassword(password ?? "", salt);
        return CryptographicOperations.FixedTimeEquals(hash, Convert.FromBase64String(user.PasswordHash)) ? user : null;
    }

    public static bool ChangePassword(string userId, string? current, string? next)
    {
        if (string.IsNullOrEmpty(next) || next!.Length < 6) return false;
        var users = Load();
        var user = users.FirstOrDefault(u => u.Id == userId);
        if (user == null) return false;
        var salt = Convert.FromBase64String(user.Salt);
        if (!CryptographicOperations.FixedTimeEquals(HashPassword(current ?? "", salt), Convert.FromBase64String(user.PasswordHash)))
            return false;
        var ns = RandomNumberGenerator.GetBytes(16);
        user.Salt = Convert.ToBase64String(ns);
        user.PasswordHash = Convert.ToBase64String(HashPassword(next!, ns));
        Save(users);
        return true;
    }

    public static UserAccount? UpdateProfile(string userId, string? name, string? email, string? role, string? avatar)
    {
        var users = Load();
        var user = users.FirstOrDefault(u => u.Id == userId);
        if (user == null) return null;
        if (!string.IsNullOrWhiteSpace(email))
        {
            var e = email.Trim();
            if (users.Any(u => u.Id != userId && u.Email.Equals(e, StringComparison.OrdinalIgnoreCase))) return null;
            user.Email = e;
        }
        if (name != null) user.Name = name.Trim();
        if (role != null) user.Role = role.Trim();
        if (avatar != null) user.Avatar = string.IsNullOrWhiteSpace(avatar) ? null : avatar;
        Save(users);
        return user;
    }

    private static byte[] HashPassword(string password, byte[] salt) =>
        Rfc2898DeriveBytes.Pbkdf2(Encoding.UTF8.GetBytes(password), salt, 100_000, HashAlgorithmName.SHA256, 32);

    // ---------- app settings (registration toggle) ----------

    public sealed class AuthSettings { public bool RegistrationOpen { get; set; } = true; }

    public static AuthSettings GetSettings()
    {
        lock (Gate)
        {
            if (!File.Exists(SettingsFile)) return new();
            try { return JsonSerializer.Deserialize<AuthSettings>(File.ReadAllText(SettingsFile)) ?? new(); }
            catch { return new(); }
        }
    }

    private static void SaveSettings(AuthSettings s)
    {
        lock (Gate) { Directory.CreateDirectory(DataDir); File.WriteAllText(SettingsFile, JsonSerializer.Serialize(s, JsonOpts)); }
    }

    public static void SetRegistrationOpen(bool open) { var s = GetSettings(); s.RegistrationOpen = open; SaveSettings(s); }

    // ---------- admin / user management ----------

    public static bool IsAdmin(UserAccount? u) => u != null && u.Role.Equals("Admin", StringComparison.OrdinalIgnoreCase);
    public static int UserCount() => Load().Count;
    public static List<object> ListUsers() => Load().OrderBy(u => u.CreatedAt).Select(Public).ToList();

    private static int AdminCount(List<UserAccount> users) => users.Count(u => u.Role.Equals("Admin", StringComparison.OrdinalIgnoreCase));

    public static (bool ok, string? error) AdminSetRole(string targetId, string? role)
    {
        role = (role ?? "").Trim();
        if (role.Length == 0) return (false, "Role is required.");
        var users = Load();
        var u = users.FirstOrDefault(x => x.Id == targetId);
        if (u == null) return (false, "User not found.");
        bool wasAdmin = u.Role.Equals("Admin", StringComparison.OrdinalIgnoreCase);
        bool willAdmin = role.Equals("Admin", StringComparison.OrdinalIgnoreCase);
        if (wasAdmin && !willAdmin && AdminCount(users) <= 1) return (false, "You can't remove the last remaining admin.");
        u.Role = role;
        Save(users);
        return (true, null);
    }

    public static (bool ok, string? error) AdminResetPassword(string targetId, string? newPassword)
    {
        if (string.IsNullOrEmpty(newPassword) || newPassword!.Length < 6) return (false, "New password must be at least 6 characters.");
        var users = Load();
        var u = users.FirstOrDefault(x => x.Id == targetId);
        if (u == null) return (false, "User not found.");
        var salt = RandomNumberGenerator.GetBytes(16);
        u.Salt = Convert.ToBase64String(salt);
        u.PasswordHash = Convert.ToBase64String(HashPassword(newPassword!, salt));
        Save(users);
        return (true, null);
    }

    public static (bool ok, string? error) AdminDelete(string adminId, string targetId)
    {
        if (adminId == targetId) return (false, "You can't delete your own account here.");
        var users = Load();
        var u = users.FirstOrDefault(x => x.Id == targetId);
        if (u == null) return (false, "User not found.");
        if (u.Role.Equals("Admin", StringComparison.OrdinalIgnoreCase) && AdminCount(users) <= 1)
            return (false, "You can't delete the last remaining admin.");
        users.Remove(u);
        Save(users);
        return (true, null);
    }

    // ---------- signed session token (userId + expiry, HMAC-signed) ----------

    private static byte[] LoadOrCreateKey()
    {
        try
        {
            Directory.CreateDirectory(DataDir);
            if (File.Exists(KeyFile)) return Convert.FromBase64String(File.ReadAllText(KeyFile));
            var k = RandomNumberGenerator.GetBytes(32);
            File.WriteAllText(KeyFile, Convert.ToBase64String(k));
            return k;
        }
        catch { return RandomNumberGenerator.GetBytes(32); }
    }

    public static string CreateToken(string userId, DateTime expiresUtc)
    {
        var payload = $"{userId}|{new DateTimeOffset(expiresUtc).ToUnixTimeSeconds()}";
        using var h = new HMACSHA256(Key);
        var sig = Convert.ToBase64String(h.ComputeHash(Encoding.UTF8.GetBytes(payload)));
        return Convert.ToBase64String(Encoding.UTF8.GetBytes(payload)) + "." + sig;
    }

    /// <summary>Returns the userId if the token is well-formed, correctly signed and unexpired.</summary>
    public static string? ValidateToken(string? token)
    {
        if (string.IsNullOrEmpty(token)) return null;
        var parts = token.Split('.');
        if (parts.Length != 2) return null;

        string payload;
        try { payload = Encoding.UTF8.GetString(Convert.FromBase64String(parts[0])); } catch { return null; }

        using var h = new HMACSHA256(Key);
        var expected = Convert.ToBase64String(h.ComputeHash(Encoding.UTF8.GetBytes(payload)));
        if (!CryptographicOperations.FixedTimeEquals(Encoding.UTF8.GetBytes(expected), Encoding.UTF8.GetBytes(parts[1])))
            return null;

        var seg = payload.Split('|');
        if (seg.Length != 2 || !long.TryParse(seg[1], out var exp)) return null;
        if (DateTimeOffset.FromUnixTimeSeconds(exp) < DateTimeOffset.UtcNow) return null;
        return seg[0];
    }

    /// <summary>Safe public projection (never exposes hash/salt).</summary>
    public static object Public(UserAccount u) => new
    {
        id = u.Id,
        username = u.Username,
        email = u.Email,
        name = u.Name,
        role = u.Role,
        avatar = u.Avatar,
    };
}
