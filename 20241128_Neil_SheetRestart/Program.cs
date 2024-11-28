using System;
using System.IO;
using System.Timers;
using System.Diagnostics;

class Program
{
    static System.Timers.Timer? timer;
    static double interval;
    static string? currentFilePath;

    static void Main()
    {
        string directoryPath = Directory.GetCurrentDirectory();
        string fileName = "prices";
        string[] extensions = { ".xlsx", ".xls", ".xlsm", ".xlsb" };

        FindFile(directoryPath, fileName, extensions);
        interval = TimerRefreshMinutes();
        StartTimer();

        // Keep the program running indefinitely
        while (true)
        {
            Thread.Sleep(1000);
        }
    }

    static void FindFile(string directoryPath, string fileName, string[] extensions)
    {
        bool fileFound = false;

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine($"Looking for '{fileName}' sheet...");
        Console.ForegroundColor = ConsoleColor.White;
        Console.WriteLine($"...in {directoryPath}");

        foreach (var extension in extensions)
        {
            string filePath = Path.Combine(directoryPath, fileName + extension);
            if (File.Exists(filePath))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("");
                Console.WriteLine($"File '{fileName}{extension}' found.");
                Console.WriteLine($"In '{directoryPath}'.");

                // Set the current file path without opening the file
                currentFilePath = filePath;

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
            Console.WriteLine("Please move this script to the same folder as the 'prices'.");
        }

        // Reset console color
        Console.ResetColor();
    }

    static void OpenFile(string filePath)
    {
        try
        {
            if (!IsFileOpen(filePath))
            {
                // Open the file
                Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                currentFilePath = filePath;
                //Console.WriteLine($"File '{filePath}' opened.");
            }
            else
            {
                Console.WriteLine($"File '{filePath}' is already open.");
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error opening file: {ex.Message}");
        }
    }

    static void CloseFile(string filePath)
    {
        try
        {
            if (IsFileOpen(filePath))
            {
                // Close the file
                foreach (var process in Process.GetProcessesByName("EXCEL"))
                {
                    if (process.MainWindowTitle.Contains(Path.GetFileNameWithoutExtension(filePath)))
                    {
                        process.Kill();
                        currentFilePath = null;
                        //Console.WriteLine($"File '{filePath}' closed.");
                        break;
                    }
                }
            }
            else
            {
                Console.WriteLine($"Cannot close the file '{filePath}' because it is not currently open.");
            }
        }
        catch (Exception ex)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"Error closing file: {ex.Message}");
        }
    }

    static bool IsFileOpen(string filePath)
    {
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            if (process.MainWindowTitle.Contains(Path.GetFileNameWithoutExtension(filePath)))
            {
                return true;
            }
        }
        return false;
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
    }

    static void OnTimedEvent(object? source, ElapsedEventArgs e)
    {
        DateTime now = DateTime.Now;
        DateTime nextRestart = now.AddMinutes(interval);
        Console.SetCursorPosition(0, Console.CursorTop);
        Console.Write(new string(' ', Console.WindowWidth)); // Clear the line
        Console.SetCursorPosition(0, Console.CursorTop);
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Restarted at {now:HH:mm}. \r Scheduled next restart in {interval} minutes at {nextRestart:HH:mm}.");

        // Close and reopen the file
        if (currentFilePath != null)
        {
            string filePath = currentFilePath; // Store the current file path
            CloseFile(filePath);
            OpenFile(filePath);
        }
    }
}
