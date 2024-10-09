using Microsoft.Win32;
using System.Runtime.Versioning;

namespace RDS;

[SupportedOSPlatform("Windows")]
internal static class Commands
{
    private const string RegistryFolder = "RDS";

    internal static void Uninstall()
    {
        Console.WriteLine("Start uninstall", ConsoleColor.Blue);

        Installer.UninstallRdp(RegistryFolder);
        Installer.UninstallAnydesk();

        Console.WriteLine("Success uninstalled", ConsoleColor.Green);
    }

    internal static void Install()
    {
        Console.WriteLine("Start install", ConsoleColor.Blue);

        Installer.InstallRdp(RegistryFolder);

        Console.WriteLine("RDP Success installed", ConsoleColor.Green);

        Console.Write("Enter Anydesk .exe file path: ", ConsoleColor.Blue);
        Installer.InstallAnydesk(Console.ReadLine());

        Console.WriteLine("Anydesk Success installed", ConsoleColor.Green);
    }

    internal static void Status()
    {
        Console.Write("RDP Status: ", ConsoleColor.Green);
        if (Installer.AlreadyInstalledRdp(RegistryFolder))
            Console.WriteLine("Installed", ConsoleColor.Green);
        else
            Console.WriteLine("Uninstalled", ConsoleColor.Red);

        Console.Write("Anydesk Status: ", ConsoleColor.Green);
        if (Installer.AlreadyInstalledAnydesk())
            Console.WriteLine("Installed", ConsoleColor.Green);
        else
            Console.WriteLine("Uninstalled", ConsoleColor.Red);
    }

    internal static void Execute(string command)
    {
        try
        {
            switch (command)
            {
                case "Uninstall": Uninstall(); break;
                case "Install": Install(); break;
                case "Status": Status(); break;
                case "Help": Help(); break;
                default: Console.WriteLine("Unknown command", ConsoleColor.Red); break;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString(), ConsoleColor.Red);
        }
    }

    internal static void Help()
    {
        var commands = $"Uninstall - remove RDS,{Environment.NewLine}Install - install RDS,{Environment.NewLine}Status - shows the status RDS,{Environment.NewLine}Help - Show help";

        Console.WriteLine(commands, ConsoleColor.White);
    }
}

[SupportedOSPlatform("Windows")]
internal static class Installer
{
    internal static void UninstallRdp(string RegistryFolder)
    {
        if (!Utils.HasAdministratorRights()) throw new Exception("You must run the program as administrator");

        Registry.ClassesRoot.DeleteSubKeyTree(RegistryFolder, false);
    }

    internal static void InstallRdp(string RegistryFolder)
    {
        if (!Utils.HasAdministratorRights()) throw new Exception("You must run the program as administrator");

        if (AlreadyInstalledRdp(RegistryFolder))
            UninstallRdp(RegistryFolder);

        var folder = Registry.ClassesRoot.CreateSubKey(RegistryFolder);
        var commandKey = folder.CreateSubKey("shell").CreateSubKey("open").CreateSubKey("command");

        folder.SetValue("", "rds:Remote Desktop Service");
        folder.SetValue("URL Protocol", "");
        commandKey.SetValue("", $@"""{Utils.AppPath}"" ""%1""");
    }

    internal static void UninstallAnydesk()
    {
        var settings = Settings.Get();
        settings.AnydeskPath = string.Empty;
        settings.Save();
    }

    internal static void InstallAnydesk(string Path)
    {
        if (!File.Exists(Path)) throw new Exception("Anydesk file not exist");

        var settings = Settings.Get();
        settings.AnydeskPath = Path;
        settings.Save();
    }

    internal static bool AlreadyInstalledRdp(string RegistryFolder) => Registry.ClassesRoot.OpenSubKey(RegistryFolder) is not null;

    internal static bool AlreadyInstalledAnydesk()
    {
        var settings = Settings.Get();

        return !string.IsNullOrEmpty(settings.AnydeskPath) && File.Exists(settings.AnydeskPath);
    }
}