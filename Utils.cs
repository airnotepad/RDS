using System.Diagnostics;
using System.Runtime.Versioning;
using System.Security.Principal;

namespace RDS;

[SupportedOSPlatform("Windows")]
internal static class Utils
{
    internal static string AppPath => Environment.ProcessPath!;
    internal static string SettingsPath => Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)!, "settings.json");

    internal static bool HasAdministratorRights()
    {
        using var user = WindowsIdentity.GetCurrent();
        try
        {
            return new WindowsPrincipal(user).IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch (UnauthorizedAccessException)
        {
            return false;
        }
    }

    internal static string GenerateFileName(RDP rdp)
    {
        var rdpFileName = $"{rdp.Address}_{rdp.Port}_{DateTime.Now.Ticks}";
        var invalid = new string(Path.GetInvalidFileNameChars()) +
                      new string(Path.GetInvalidPathChars());
        return invalid.Aggregate(rdpFileName, (current, c) => current.Replace(c.ToString(), ""));
    }
}
