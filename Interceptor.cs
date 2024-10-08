using System.Diagnostics;
using System.Runtime.Versioning;

namespace RDS;

[SupportedOSPlatform("Windows")]
internal class Interceptor
{
    internal static void Execute(IEnumerable<string> args)
    {
        var parameters = Patameters.Parse(args);
        RunRDP(parameters.RDP);
    }

    private static void RunRDP(RDP rdp)
    {
        var rdpFile = Path.Combine(Path.GetTempPath(), $"{Utils.GenerateFileName(rdp)}.rdp");
        File.WriteAllText(rdpFile, rdp.ToRdpConfig().ToString());

        string admin = rdp.IsAdministrativePurposes == true ? " /admin " : "";
        new Process
        {
            StartInfo =
            {
                FileName = "mstsc.exe",
                Arguments = $"{admin} \"{rdpFile}\""
            },
        }.Start();

        Task.Run(() =>
        {
            var attempts = 30;
            while (attempts > 0)
            {
                Task.Delay(1000).Wait();
                if (File.Exists(rdpFile))
                {
                    File.Delete(rdpFile);
                    return;
                }
                attempts--;
            }
        }).Wait();
    }
}