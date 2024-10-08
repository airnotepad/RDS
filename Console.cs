namespace RDS;

internal static class Console
{
    private static readonly ConsoleColor defaultColor = ConsoleColor.White;
    internal static void WriteLine(string? message = null, ConsoleColor color = ConsoleColor.White) 
    { 
        System.Console.ForegroundColor = color;
        System.Console.WriteLine(message);
        System.Console.ForegroundColor = defaultColor;
    }
    internal static void Write(string message, ConsoleColor color)
    {
        System.Console.ForegroundColor = color;
        System.Console.Write(message);
        System.Console.ForegroundColor = defaultColor;
    }
    internal static string ReadLine() => System.Console.ReadLine() ?? "";
}
