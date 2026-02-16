using NanoXLSX;
using NanoXLSX.Styles;
using System.Collections.Generic;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of cell ranges
    /// </summary>
    public class CellRanges
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("CellRanges.xlsx", "Sheet1");                              // Create new workbook
            List<object> values = new List<object>() { "Header1", "Header2", "Header3" };               // Create a List of values
            workbook.CurrentWorksheet.AddCellRange(values, "A1:C1");                                    // Add cell range

            values = new List<object>() { "Cell A2", "Cell B2", "Cell C2" };                            // Create a List of values
            workbook.CurrentWorksheet.AddCellRange(values, new Address("A2"), new Address("C2"));       // Add cell range (using addresses)

            values = new List<object>() { "Cell A3", "Cell B3", "Cell C3" };                            // Create a List of values
            workbook.CurrentWorksheet.AddCellRange(values, "A3:C3", BasicStyles.Bold);                  // Add cell range (using a style)

            workbook.Save();                                                                            // Save the workbook
        }

        private CellRanges()
        {
            // Do not instantiate
        }
    }
}
