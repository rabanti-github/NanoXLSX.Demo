using System;
using System.Collections.Generic;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of hiding rows and columns
    /// </summary>
    public class HidingRowsAndColumns
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("worksheet1");                                             // Create new workbook without worksheet
            Worksheet ws = workbook.CurrentWorksheet;                                                   // Create reference (shortening)
            List<object> values = new List<object>() { "Cell A1", "Cell B1", "Cell C1", "Cell D1" };    // Create a List of values
            ws.AddCellRange(values, "A1:D1");                                                           // Insert cell range
            values = new List<object>() { "Cell A2", "Cell B2", "Cell C2", "Cell D2" };                 // Create a List of values
            ws.AddCellRange(values, "A2:D2");                                                           // Insert cell range
            values = new List<object>() { "Cell A3", "Cell B3", "Cell C3", "Cell D3" };                 // Create a List of values
            ws.AddCellRange(values, "A3:D3");                                                           // Insert cell range
            ws.AddHiddenColumn("C");                                                                    // Hide column C
            ws.AddHiddenRow(1);                                                                         // Hider row 2 (zero-based: 1)
            workbook.SaveAs("HidingRowsAndColumns.xlsx");                                               // Save the workbook
        }

        private HidingRowsAndColumns()
        {
            // Do not instantiate
        }
    }
}