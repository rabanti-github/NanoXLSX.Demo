using NanoXLSX.Styles;
using System;
using System.Collections.Generic;

namespace NanoXLSX.Demo.UseCases
{
    public class ColumnWidthsRowHeights
    {

        /// <summary>
        /// This demo shows the usage of column widths and row heights
        /// </summary>
        public static void Run()
        {
            Workbook workbook = new Workbook("ColumnWidthsRowHeights.xlsx", "Sheet1");                      // Create new workbook
            workbook.CurrentWorksheet.AddCell("test", "A1");                                                // Write some data
            workbook.CurrentWorksheet.AddCell("test", "B2");
            workbook.CurrentWorksheet.AddCell("test", "C3");
            workbook.CurrentWorksheet.SetColumnWidth(0, 20f);                                               // Set column width
            workbook.CurrentWorksheet.SetColumnWidth(1, 15f);                                               // Set column width
            workbook.CurrentWorksheet.SetColumnWidth(2, 25f);                                               // Set column width
            workbook.CurrentWorksheet.SetRowHeight(0, 20);                                                  // Set row height
            workbook.CurrentWorksheet.SetRowHeight(1, 30);                                                  // Set row height

            workbook.Save();                                                                                // Save the workbook
        }

        private ColumnWidthsRowHeights()
        {
            // Do not instantiate
        }
    }
}