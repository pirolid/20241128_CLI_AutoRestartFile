using System;
using System.IO;

class Program
{
    static void Main()
    {
        string directoryPath = Directory.GetCurrentDirectory();
        string fileName = "prices";
        string[] extensions = { ".xlsx", ".xls", ".xlsm", ".xlsb" };

        FindFile(directoryPath, fileName, extensions);
        double interval = TimerRefreshMinutes();
        Console.WriteLine($"Resetting file every {interval} min.");
        ExitProgram();
    }

    static void FindFile(string directoryPath, string fileName, string[] extensions)
    {
        bool fileFound = false;

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"Looking for '{fileName}' sheet...");
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine($"Search current folder: {directoryPath}");

        foreach (var extension in extensions)
        {
            string filePath = Path.Combine(directoryPath, fileName + extension);
            if (File.Exists(filePath))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("");
                Console.WriteLine($"File '{fileName}{extension}' found.");
                Console.WriteLine($"In directory '{directoryPath}'.");
                Console.WriteLine("");
                fileFound = true;
                break;
            }
        }

        if (!fileFound)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"File '{fileName}' not found.");
            Console.WriteLine("");
            Console.WriteLine("Please move this script to the same folder as the 'prices' Excel sheet.");
        }

        // Reset console color
        Console.ResetColor();
    }

    static double TimerRefreshMinutes()
    {
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.Write("Enter restart timer in minutes: ");
        if (double.TryParse(Console.ReadLine(), out double interval))
        {
            Console.ForegroundColor = ConsoleColor.Green;
            return interval;
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Invalid input. Please enter a valid number.");
            return TimerRefreshMinutes(); // Recursively ask for input again
        }
    }

    static void ExitProgram()
    {
        // Keep the console window open for troubleshooting
        Console.ForegroundColor = ConsoleColor.Magenta;
        Console.WriteLine("Press Escape key to exit...");
        while (Console.ReadKey(true).Key != ConsoleKey.Escape)
        {
            // Wait for the Escape key to be pressed
        }
        Environment.Exit(0); // Exit the program immediately
    }
}