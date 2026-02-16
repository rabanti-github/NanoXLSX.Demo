using System;
using System.Collections.Generic;
using System.Linq;
using UC = PicoXLSX.Demo.UseCases;

namespace PicoXLSX.Demo
{
    /// <summary>
    /// Main class to run PicoXLSX demo use cases
    /// </summary>
    public class Program
    {
        private static readonly Dictionary<string, (Action Run, string Description)> UseCases = new()
        {
            { "BasicDemo", (UC.BasicDemo.Run, "Create a basic workbook with text, number, and datetime values") },
            { "Shortener", (UC.Shortener.Run, "Demonstrate shortened syntax for writing cells") },
            { "CellAndWorksheetSelection", (UC.CellAndWorksheetSelection.Run, "Usage of cell and worksheet selection") },
            { "Stream", (UC.Stream.Run, "Save workbooks to memory streams and file streams") },
            { "Async", (() => UC.Async.Run().GetAwaiter().GetResult(), "Asynchronous file saving") },
            { "AddNextCell", (UC.AddnextCell.Run, "Using AddNextCell for sequential cell addition") },
            { "DataTypes", (UC.DataTypes.Run, "Working with various data types") },
            { "CellDirectionsAndValues", (UC.CellDirectionsAndValues.Run, "Cell directions and value handling") },
            { "ColumnWidthsRowHeights", (UC.ColumnWidthsRowHeights.Run, "Setting column widths and row heights") },
            { "StyleBasics", (UC.StyleBasics.Run, "Basic usage of styles") },
            { "StyleAppending", (UC.StyleAppending.Run, "Style appending and method chaining") },
            { "ActiveAndSetStyle", (UC.ActiveAndSetStyle.Run, "Applying styles to cells and ranges") },
            { "CellRanges", (UC.CellRanges.Run, "Working with cell ranges") },
            { "Metadata", (UC.Metadata.Run, "Assigning workbook metadata") },
            { "MergeCells", (UC.MergeCells.Run, "Merging cells") },
            { "ProtectionAndPasswords", (UC.ProtectionAndPasswords.Run, "Sheet protection and password protection") },
            { "HidingRowsAndColumns", (UC.HidingRowsAndColumns.Run, "Hiding rows and columns") },
            { "AutoFilter", (UC.AutoFilter.Run, "Auto-filtering") },
            { "SanitizingWorksheetNames", (UC.SanitizingWorksheetNames.Run, "Worksheet name sanitization") },
            { "Formulas", (UC.Formulas.Run, "Excel formulas usage") },
            { "PaneSplitAndFreeze", (UC.PaneSplitAndFreeze.Run, "Worksheet splitting and freezing panes") },
            { "HidingWorkbooksAndWorksheets", (UC.HidingWorkbooksAndWorksheets.Run, "Hiding workbooks and worksheets") },
            { "InlineFormattingBasics", (UC.InlineFormattingBasics.Run, "Inline formatting with multiple text runs") },
            { "InlineFormattingStyles", (UC.InlineFormattingStyles.Run, "Styled inline formatting (bold, italic, colors)") },
            { "InlineFormattingPhonetic", (UC.InlineFormattingPhonetic.Run, "Phonetic runs for East Asian languages") }
        };

        /// <summary>
        /// Main entry point
        /// </summary>
        /// <param name="args">
        /// Optional arguments:
        /// - No args: Shows menu to select use cases
        /// - "all": Runs all use cases
        /// - Comma-separated use case names or numbers: Runs specified use cases (e.g., "BasicDemo,Stream" or "1,2")
        /// </param>
        static void Main(string[] args)
        {
            Console.WriteLine("=================================================");
            Console.WriteLine("  PicoXLSX Demo - Use Case Examples (v4.x)");
            Console.WriteLine("=================================================\n");

            if (args.Length == 0)
            {
                RunInteractive();
            }
            else if (args[0].Equals("all", StringComparison.OrdinalIgnoreCase))
            {
                RunAllUseCases();
            }
            else
            {
                RunSpecificUseCases(args[0]);
            }

            Console.WriteLine("\n=================================================");
            Console.WriteLine("  Demo execution completed!");
            Console.WriteLine("=================================================");
        }

        /// <summary>
        /// Runs all use cases sequentially
        /// </summary>
        private static void RunAllUseCases()
        {
            Console.WriteLine("Running ALL use cases...\n");
            int count = 1;
            foreach (var useCase in UseCases)
            {
                RunUseCase(useCase.Key, count++, UseCases.Count);
            }
        }

        /// <summary>
        /// Runs specific use cases by name or number
        /// </summary>
        private static void RunSpecificUseCases(string selection)
        {
            var selections = selection.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                      .Select(s => s.Trim())
                                      .ToArray();
            var useCaseList = UseCases.Keys.ToList();

            foreach (var sel in selections)
            {
                if (int.TryParse(sel, out int index) && index > 0 && index <= useCaseList.Count)
                {
                    var useCaseName = useCaseList[index - 1];
                    RunUseCase(useCaseName, index, UseCases.Count);
                }
                else if (UseCases.ContainsKey(sel))
                {
                    int idx = useCaseList.IndexOf(sel) + 1;
                    RunUseCase(sel, idx, UseCases.Count);
                }
                else
                {
                    Console.WriteLine($"[WARNING] Use case '{sel}' not found. Skipping...\n");
                }
            }
        }

        /// <summary>
        /// Interactive menu for selecting use cases
        /// </summary>
        private static void RunInteractive()
        {
            while (true)
            {
                DisplayMenu();
                Console.Write("\nEnter your choice (0 to exit, 'all' for all use cases, or comma-separated numbers/names): ");
                string input = Console.ReadLine()?.Trim();

                if (string.IsNullOrEmpty(input) || input == "0")
                {
                    Console.WriteLine("Exiting...");
                    break;
                }

                if (input.Equals("all", StringComparison.OrdinalIgnoreCase))
                {
                    RunAllUseCases();
                    break;
                }

                RunSpecificUseCases(input);

                Console.Write("\nPress any key to continue or 'Q' to quit...");
                if (Console.ReadKey().Key == ConsoleKey.Q)
                {
                    Console.WriteLine("\nExiting...");
                    break;
                }
                Console.WriteLine("\n");
            }
        }

        /// <summary>
        /// Displays the menu of available use cases
        /// </summary>
        private static void DisplayMenu()
        {
            Console.WriteLine("\nAvailable Use Cases:");
            Console.WriteLine("-------------------");
            int index = 1;
            foreach (var useCase in UseCases)
            {
                Console.WriteLine($"  {index,2}. {useCase.Key,-35} - {useCase.Value.Description}");
                index++;
            }
            Console.WriteLine("   0. Exit");
        }

        /// <summary>
        /// Runs a single use case
        /// </summary>
        private static void RunUseCase(string name, int current, int total)
        {
            if (!UseCases.ContainsKey(name))
            {
                return;
            }

            Console.WriteLine($"[{current}/{total}] Running: {name}");
            Console.WriteLine($"Description: {UseCases[name].Description}");
            Console.WriteLine(new string('-', 80));

            try
            {
                UseCases[name].Run();
                Console.WriteLine($"✓ {name} completed successfully\n");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"✗ {name} failed with error:");
                Console.WriteLine($"  {ex.Message}\n");
                Console.WriteLine($"Stack trace:\n{ex.StackTrace}\n");
            }
        }

        private Program()
        {
            // Do not instantiate
        }
    }
}
