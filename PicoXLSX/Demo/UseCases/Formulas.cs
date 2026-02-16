using NanoXLSX;
using NanoXLSX.Styles;
using System.Collections.Generic;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of basic Excel formulas
    /// </summary>
    public class Formulas
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("Formulas.xlsx", "sheet1");                                // Create a new workbook
            List<object> numbers = new List<object> { 1.15d, 2.225d, 13.8d, 15d, 15.1d, 17.22d, 22d, 107.5d, 128d }; // Create a list of numbers
            List<object> texts = new List<object>() { "value 1", "value 2", "value 3", "value 4", "value 5", "value 6", "value 7", "value 8", "value 9" }; // Create a list of strings (for vlookup)
            workbook.WS.Value("Numbers", BasicStyles.Bold);                                             // Add a header with a basic style
            workbook.WS.Value("Values", BasicStyles.Bold);                                              // Add a header with a basic style
            workbook.WS.Value("Formula type", BasicStyles.Bold);                                        // Add a header with a basic style
            workbook.WS.Value("Formula value", BasicStyles.Bold);                                       // Add a header with a basic style
            workbook.WS.Value("(See also worksheet2)");                                                 // Add a note
            workbook.CurrentWorksheet.AddCellRange(numbers, "A2:A10");                                  // Add the numbers as range
            workbook.CurrentWorksheet.AddCellRange(texts, "B2:B10");                                    // Add the values as range

            workbook.CurrentWorksheet.SetCurrentCellAddress("D2");                                      // Set the "cursor" to D2
            Cell c;                                                                                     // Create an empty cell object (reusable)
            c = BasicFormulas.Average(new Range("A2:A10"));                                             // Define an average formula
            workbook.CurrentWorksheet.AddCell("Average", "C2");                                         // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D2");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Ceil(new Address("A2"), 0);                                               // Define a ceil formula
            workbook.CurrentWorksheet.AddCell("Ceil", "C3");                                            // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D3");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Floor(new Address("A2"), 0);                                              // Define a floor formula
            workbook.CurrentWorksheet.AddCell("Floor", "C4");                                           // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D4");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Round(new Address("A3"), 1);                                              // Define a round formula with one digit after the comma
            workbook.CurrentWorksheet.AddCell("Round", "C5");                                           // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D5");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Max(new Range("A2:A10"));                                                 // Define a max formula
            workbook.CurrentWorksheet.AddCell("Max", "C6");                                             // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D6");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Min(new Range("A2:A10"));                                                 // Define a min formula
            workbook.CurrentWorksheet.AddCell("Min", "C7");                                             // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D7");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Median(new Range("A2:A10"));                                              // Define a median formula
            workbook.CurrentWorksheet.AddCell("Median", "C8");                                          // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D8");                                                 // Add the formula to the worksheet

            c = BasicFormulas.Sum(new Range("A2:A10"));                                                 // Define a sum formula
            workbook.CurrentWorksheet.AddCell("Sum", "C9");                                             // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D9");                                                 // Add the formula to the worksheet

            c = BasicFormulas.VLookup(13.8d, new Range("A2:B10"), 2, true);                             // Define a vlookup formula (look for the value of the number 13.8)
            workbook.CurrentWorksheet.AddCell("Vlookup", "C10");                                        // Add the description of the formula to the worksheet
            workbook.CurrentWorksheet.AddCell(c, "D10");                                                // Add the formula to the worksheet

            workbook.AddWorksheet("sheet2");                                                            // Create a new worksheet
            c = BasicFormulas.VLookup(workbook.Worksheets[0], new Address("B4"), workbook.Worksheets[0], new Range("B2:C10"), 2, true); // Define a vlookup formula in worksheet1 (look for the text right of the (value of) cell B4)
            workbook.WS.Value(c);                                                                       // Add the formula to the worksheet

            c = BasicFormulas.Median(workbook.Worksheets[0], new Range("A2:A10"));                      // Define a median formula in worksheet1
            workbook.WS.Value(c);                                                                       // Add the formula to the worksheet

            workbook.Save();                                                                            // Save the workbook
        }

        private Formulas()
        {
            // Do not instantiate
        }
    }
}
