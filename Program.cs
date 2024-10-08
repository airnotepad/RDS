using System.Runtime.Versioning;

namespace RDS;

[SupportedOSPlatform("Windows")]
internal class Program
{
    static void Main(string[] args)
    {
        if (args.Length > 0)
        {
            Interceptor.Execute(args);
            return;
        }

        Commands.Help();
        Console.WriteLine();
        Console.WriteLine("Enter any command ... or Enter for exit", ConsoleColor.Blue);

        Run();
    }

    private static void Run()
    {
        string command = string.Empty;

        do
        {
            Console.WriteLine();
            Console.Write("Your command: ", ConsoleColor.Blue);
            command = Console.ReadLine();
            Console.WriteLine();
            Commands.Execute(command);
        } while (!string.IsNullOrEmpty(command));
    }
}
