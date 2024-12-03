﻿using System;
using System.IO;
using System.Timers;
using System.Diagnostics;
using OfficeOpenXml; // Install-Package EPPlus

class Program
{
    static System.Timers.Timer? timer;
    static double interval;
    static string? currentFilePath;

    static void Main()
    {
        // Ensure EPPlus is licensed
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Add this line

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

        Console.ForegroundColor = ConsoleColor.Blue;
        CenterText($"Looking for '{fileName}' sheet...");
        Console.ForegroundColor = ConsoleColor.White;
        CenterText($"...in {directoryPath}");

        foreach (var extension in extensions)
        {
            string filePath = Path.Combine(directoryPath, fileName + extension);
            if (File.Exists(filePath))
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("");
                CenterText($"File '{fileName}{extension}' found.");
                //Console.WriteLine($"In '{directoryPath}'.");

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
            CenterText($"File '{fileName}' not found.");
            Console.WriteLine("");
            CenterText("Please ensure this script is placed in the same folder as the 'prices' file.");
            Console.WriteLine("");
            CenterText("Press any key to exit...");
            Console.ResetColor();
            Console.ReadKey(); // Wait for the user to press any key
            Environment.Exit(0); // Exit the program
        }

        // Reset console color
        Console.ResetColor();
    }

    static void OpenFile(string filePath)
    {
        int retryCount = 3;
        int delay = 5000; // 5 seconds

        for (int i = 0; i < retryCount; i++)
        {
            try
            {
                if (!IsFileOpen(filePath))
                {
                    // Open the file
                    Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
                    currentFilePath = filePath;
                    Console.WriteLine("");
                    Console.WriteLine($"File '{filePath}' opened.");
                    return;
                }
                else
                {
                    Console.WriteLine($"File already open '{filePath}'.");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error opening file: {ex.Message}");
                if (i < retryCount - 1)
                {
                    Console.WriteLine($"Retrying in {delay / 1000} seconds...");
                    Thread.Sleep(delay);
                }
            }
        }
        Console.WriteLine("Failed to open the file after multiple attempts.");
    }

    static void CloseFile(string filePath)
    {
        int retryCount = 3;
        int delay = 5000; // 5 seconds

        for (int i = 0; i < retryCount; i++)
        {
            try
            {
                if (IsFileOpen(filePath))
                {
                    // Close the file using EPPlus
                    using (var package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        package.Save(); // Save any changes
                        currentFilePath = null;
                        Console.WriteLine($"File '{filePath}' closed.");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine($"File already closed '{filePath}'.");
                    return;
                }
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error closing file: {ex.Message}");
                if (i < retryCount - 1)
                {
                    Console.WriteLine($"Retrying in {delay / 1000} seconds...");
                    Thread.Sleep(delay);
                }
            }
        }
        Console.WriteLine("Failed to close the file after multiple attempts.");
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
        CenterText("Enter restart timer in minutes: ", true);
        Console.Write("");
        if (double.TryParse(Console.ReadLine(), out double interval))
        {
            Console.ForegroundColor = ConsoleColor.Green;
            return interval;
        }
        else
        {
            Console.ForegroundColor = ConsoleColor.Red;
            CenterText("Invalid input. Please enter a valid number.");
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
        Console.WriteLine("");
        CenterText($"Timer started. Restarting every {interval} minutes.");
        StartCountdown(interval);
    }

    static void StartCountdown(double minutes)
    {
        int totalSeconds = (int)(minutes * 60);
        while (totalSeconds > 0)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.SetCursorPosition(0, Console.CursorTop);
            CenterText($"Time until next restart: {totalSeconds / 60:D2}:{totalSeconds % 60:D2}  ", true);
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
        CenterText($"Restarted at {now:HH:mm}. \r Scheduled next restart in {interval} minutes at {nextRestart:HH:mm}.");
        Console.Write("");

        // Close and reopen the file
        if (currentFilePath != null)
        {
            string filePath = currentFilePath; // Store the current file path
            CloseFile(filePath);
            OpenFile(filePath);
        }

        StartCountdown(interval); // Restart the countdown after each timer event
    }

    static void CenterText(string text, bool inputPrompt = false)
    {
        int windowWidth = Console.WindowWidth;
        int textLength = text.Length;
        int spaces = Math.Max((windowWidth - textLength) / 2, 0); // Ensure spaces is non-negative

        if (textLength > windowWidth)
        {
            // If the text is longer than the window width, print it without centering
            Console.WriteLine(text);
        }
        else
        {
            if (inputPrompt)
            {
                Console.SetCursorPosition(spaces, Console.CursorTop);
                Console.Write(text);
            }
            else
            {
                Console.WriteLine(new string(' ', spaces) + text);
            }
        }
    }
}
