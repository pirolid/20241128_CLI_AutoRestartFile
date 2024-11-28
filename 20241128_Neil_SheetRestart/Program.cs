using System;
using System.IO;
using System.Timers;

class Program
{
    static System.Timers.Timer? timer;
    static double interval;

    static void Main()
    {
        string directoryPath = Directory.GetCurrentDirectory();
        string fileName = "prices";
        string[] extensions = { ".xlsx", ".xls", ".xlsm", ".xlsb" };

        FindFile(directoryPath, fileName, extensions);
        interval = TimerRefreshMinutes();
        StartTimer();
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

    static void StartTimer()
    {
        timer = new System.Timers.Timer(interval * 60 * 1000); // Convert minutes to milliseconds
        timer.Elapsed += OnTimedEvent;
        timer.AutoReset = true;
        timer.Enabled = true;

        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Timer started. Restarting every {interval} minutes.");

        // Start the countdown
        StartCountdown(interval);
    }

    static void StartCountdown(double minutes)
    {
        int totalSeconds = (int)(minutes * 60);
        while (totalSeconds > 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write($"Time until next restart: {totalSeconds / 60:D2}:{totalSeconds % 60:D2}  ");
            Thread.Sleep(1000);
            totalSeconds--;
        }
        Console.SetCursorPosition(0, Console.CursorTop);
        Console.Write(new string(' ', Console.WindowWidth)); // Clear the line
        Console.SetCursorPosition(0, Console.CursorTop);
    }

    static void OnTimedEvent(object? source, ElapsedEventArgs e)
    {
        DateTime now = DateTime.Now;
        DateTime nextRestart = now.AddMinutes(interval);
        Console.SetCursorPosition(0, Console.CursorTop);
        Console.Write(new string(' ', Console.WindowWidth)); // Clear the line
        Console.SetCursorPosition(0, Console.CursorTop);
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Restarted at {now:HH:mm}. Scheduled next restart in {interval} minutes at {nextRestart:HH:mm}.");
        StartCountdown(interval); // Restart the countdown after each timer event
    }

    static void ExitProgram()
    {
        // Keep the console window open for troubleshooting
        Console.ForegroundColor = ConsoleColor.Magenta;
        Console.WriteLine("");
        Console.WriteLine("Press Escape key to exit...");
        while (Console.ReadKey(true).Key != ConsoleKey.Escape)
        {
            // Wait for the Escape key to be pressed
        }
        timer?.Stop();
        timer?.Dispose();
        Environment.Exit(0); // Exit the program immediately
    }
}
