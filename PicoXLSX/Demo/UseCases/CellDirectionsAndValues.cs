using NanoXLSX;
using System;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of flipped direction when using AddNextCell, reading of the current cell address, and reading of cell values
    /// </summary>
    public class CellDirectionsAndValues
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("CellDirectionsAndValues.xlsx", "Sheet1");         // Create new workbook
            workbook.CurrentWorksheet.CurrentCellDirection = Worksheet.CellDirection.RowToRow;  // Change the cell direction
            workbook.CurrentWorksheet.AddNextCell(1);                                           // Add cell A1
            workbook.CurrentWorksheet.AddNextCell(2);                                           // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(3);                                           // Add cell A3
            workbook.CurrentWorksheet.AddNextCell(4);                                           // Add cell A4
            int row = workbook.CurrentWorksheet.GetCurrentRowNumber();                          // Get the row number (will be 4 = row 5)
            int col = workbook.CurrentWorksheet.GetCurrentColumnNumber();                       // Get the column number (will be 0 = column A)
            workbook.CurrentWorksheet.AddNextCell("This cell has the row number " + (row + 1) + " and column number " + (col + 1));
            workbook.CurrentWorksheet.GoToNextColumn();                                         // Go to Column B
            workbook.CurrentWorksheet.AddNextCell("A");                                         // Add cell B1
            workbook.CurrentWorksheet.AddNextCell("B");                                         // Add cell B2
            workbook.CurrentWorksheet.AddNextCell("C");                                         // Add cell B3
            workbook.CurrentWorksheet.AddNextCell("D");                                         // Add cell B4
            workbook.CurrentWorksheet.RemoveCell("A2");                                         // Delete cell A2
            workbook.CurrentWorksheet.RemoveCell(1, 1);                                         // Delete cell B2
            workbook.CurrentWorksheet.GoToNextRow(3);                                           // Move 3 rows down
            object value = workbook.CurrentWorksheet.GetCell(1, 2).Value;                       // Gets the value of cell B3
            workbook.CurrentWorksheet.AddNextCell("Value of B3 is: " + value);
            workbook.CurrentWorksheet.CurrentCellDirection = Worksheet.CellDirection.Disabled;  // Disable automatic cell addressing
            workbook.CurrentWorksheet.AddCell("Text A", 3, 0);                                  // Add manually placed value
            workbook.CurrentWorksheet.AddCell("Text B", 4, 1);                                  // Add manually placed value
            workbook.CurrentWorksheet.AddCell("Text C", 3, 2);                                  // Add manually placed value
            workbook.Save();                                                                    // Save the workbook
        }

        private CellDirectionsAndValues()
        {
            // Do not instantiate
        }
    }
}
