using NanoXLSX.Extensions;
using System;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows how to read cells with inline formatting from an Excel file
    /// </summary>
    public class InlineFormattingRead
    {
        public const string FORMATTING_FILE = "InlineFormattingStyles.xlsx";                        // File created by InlineFormattingStyles demo

        public static void Run()
        {
            // First, ensure the file exists by running InlineFormattingStyles
            Console.WriteLine("Creating sample file with inline formatting...");
            InlineFormattingStyles.Run();                                                           // Create the sample file

            // Now load and read the formatted text
            Console.WriteLine("\nReading formatted text from file...\n");
            Workbook loadedWorkbook = WorkbookReader.Load(FORMATTING_FILE);                         // Load the workbook

            // Read cell A1 (Bold and italic example)
            FormattedText formatted1 = loadedWorkbook.CurrentWorksheet.GetCell(0, 0).Value as FormattedText;
            if (formatted1 != null)
            {
                Console.WriteLine("Cell A1:");
                Console.WriteLine($"  Plain text: {formatted1.PlainText}");
                Console.WriteLine($"  Number of runs: {formatted1.Runs.Count}");
                foreach (var run in formatted1.Runs)
                {
                    Console.WriteLine($"    Run: '{run.Text}' - Bold: {run.FontStyle?.Bold}, Italic: {run.FontStyle?.Italic}");
                }
            }

            // Read cell A2 (Color and size example)
            FormattedText formatted2 = loadedWorkbook.CurrentWorksheet.GetCell(0, 1).Value as FormattedText;
            if (formatted2 != null)
            {
                Console.WriteLine("\nCell A2:");
                Console.WriteLine($"  Plain text: {formatted2.PlainText}");
                foreach (var run in formatted2.Runs)
                {
                    Console.WriteLine($"    Run: '{run.Text}' - Size: {run.FontStyle?.Size}, Color: {run.FontStyle?.ColorValue}");
                }
            }

            // Read cell A3 (InlineStyleBuilder example)
            FormattedText formatted3 = loadedWorkbook.CurrentWorksheet.GetCell(0, 2).Value as FormattedText;
            if (formatted3 != null)
            {
                Console.WriteLine("\nCell A3:");
                Console.WriteLine($"  Plain text: {formatted3.PlainText}");
                Console.WriteLine($"  Total runs: {formatted3.Runs.Count}");
            }

            // Read cell A4 (Complex formatting example)
            FormattedText formatted4 = loadedWorkbook.CurrentWorksheet.GetCell(0, 3).Value as FormattedText;
            if (formatted4 != null)
            {
                Console.WriteLine("\nCell A4:");
                Console.WriteLine($"  Plain text: {formatted4.PlainText}");
                Console.WriteLine($"  Runs breakdown:");
                for (int i = 0; i < formatted4.Runs.Count; i++)
                {
                    var run = formatted4.Runs[i];
                    Console.WriteLine($"    Run {i + 1}: '{run.Text}'");
                    if (run.FontStyle != null)
                    {
                        Console.WriteLine($"      - Bold: {run.FontStyle.Bold}, Italic: {run.FontStyle.Italic}");
                        Console.WriteLine($"      - Underline: {run.FontStyle.Underline}, Size: {run.FontStyle.Size}");
                        Console.WriteLine($"      - Color: {run.FontStyle.ColorValue}");
                    }
                }
            }

            Console.WriteLine("\nFormatted text reading completed!");
        }

        private InlineFormattingRead()
        {
            // Do not instantiate
        }
    }
}
