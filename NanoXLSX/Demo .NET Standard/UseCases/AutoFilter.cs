using System;
using System.Collections.Generic;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of auto-filters
    /// </summary>
    public class AutoFilter
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("sheet1");                                                 // Create new workbook without worksheet
            Worksheet ws = workbook.CurrentWorksheet;                                                   // Create reference (shortening)
            List<object> values = new List<object>() { "Cell A1", "Cell B1", "Cell C1", "Cell D1" };    // Create a List of values
            ws.AddCellRange(values, "A1:D1");                                                           // Insert cell range
            values = new List<object>() { "Cell A2", "Cell B2", "Cell C2", "Cell D2" };                 // Create a List of values
            ws.AddCellRange(values, "A2:D2");                                                           // Insert cell range
            values = new List<object>() { "Cell A3", "Cell B3", "Cell C3", "Cell D3" };                 // Create a List of values
            ws.AddCellRange(values, "A3:D3");                                                           // Insert cell range
            ws.SetAutoFilter(1, 3);                                                                     // Set auto-filter for column B to D
            workbook.SaveAs("AutoFilter.xlsx");                                                         // Save the workbook
        }

        private AutoFilter()
        {
            // Do not instantiate
        }
    }
}