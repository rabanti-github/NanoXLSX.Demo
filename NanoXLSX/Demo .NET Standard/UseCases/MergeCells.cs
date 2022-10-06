using System;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of merging cells
    /// </summary>
    public class MergeCells
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("MergeCells.xlsx", "Sheet1");                              // Create new workbook
            workbook.CurrentWorksheet.AddNextCell("Merged1");                                           // Add cell A1
            workbook.CurrentWorksheet.MergeCells("A1:C1");                                              // Merge cells from A1 to C1
            workbook.CurrentWorksheet.GoToNextRow();                                                    // Go to next row
            workbook.CurrentWorksheet.AddNextCell(false);                                               // Add cell A2
            workbook.CurrentWorksheet.MergeCells("A2:D2");                                              // Merge cells from A2 to D1
            workbook.CurrentWorksheet.GoToNextRow();                                                    // Go to next row
            workbook.CurrentWorksheet.AddNextCell("22.2d");                                             // Add cell A3
            workbook.CurrentWorksheet.MergeCells("A3:E4");                                              // Merge cells from A3 to E4
            workbook.Save();                                                                            // Save the workbook
        }

        private MergeCells()
        {
            // Do not instantiate
        }
    }
}