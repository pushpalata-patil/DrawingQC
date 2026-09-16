namespace DrawingQC.Web;

/// <summary>
/// Shared constants for the DrawingQC Local Agent — the small Windows companion app that runs on
/// the engineer's PC and drives their own AutoCAD / Plant 3D and Microsoft Word. The hosted web
/// page talks to it directly at http://127.0.0.1:{DefaultPort}.
/// </summary>
public static class AgentInfo
{
    public const int DefaultPort = 5081;
    public const string Name = "DrawingQC Local Agent";
    public const string Version = "1.0.0";
}
