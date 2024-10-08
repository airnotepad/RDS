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

        Installer.Uninstall(RegistryFolder);

        Console.WriteLine("Success uninstalled", ConsoleColor.Green);
    }

    internal static void Install()
    {
        Console.WriteLine("Start install", ConsoleColor.Blue);

        Installer.Install(RegistryFolder);

        Console.WriteLine("Success installed", ConsoleColor.Green);
    }

    internal static void Status()
    {
        Console.Write("Status: ", ConsoleColor.Green);
        if (Installer.AlreadyInstalled(RegistryFolder))
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
    internal static void Uninstall(string RegistryFolder)
    {
        if (!Utils.HasAdministratorRights()) throw new Exception("You must run the program as administrator");

        Registry.ClassesRoot.DeleteSubKeyTree(RegistryFolder, false);
    }

    internal static void Install(string RegistryFolder)
    {
        if (!Utils.HasAdministratorRights()) throw new Exception("You must run the program as administrator");

        if (AlreadyInstalled(RegistryFolder))
            Uninstall(RegistryFolder);

        var folder = Registry.ClassesRoot.CreateSubKey(RegistryFolder);
        var commandKey = folder.CreateSubKey("shell").CreateSubKey("open").CreateSubKey("command");

        folder.SetValue("", "rds:Remote Desktop Service");
        folder.SetValue("URL Protocol", "");
        commandKey.SetValue("", $@"""{Utils.AppPath}"" ""%1""");
    }

    internal static bool AlreadyInstalled(string RegistryFolder) => Registry.ClassesRoot.OpenSubKey(RegistryFolder) is not null;
}