using System;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of splitting and freezing a worksheet into several panes
    /// </summary>
    public class PaneSplitAndFreeze
    {
        public static void Run()
        {
            Workbook wb = new Workbook("PaneSplitAndFreeze.xlsx", "splitXchars");                                   // Create a new workbook
            wb.CurrentWorksheet.SetVerticalSplit(30f, new Address("D1"), Worksheet.WorksheetPane.topRight);         // Split worksheet vertically by characters
            wb.AddWorksheet("SplitXcols");                                                                          // Create new worksheet
            wb.CurrentWorksheet.SetVerticalSplit(4, false, new Address("E1"), Worksheet.WorksheetPane.topRight);    // Split worksheet vertically by columns
            wb.CurrentWorksheet.SetColumnWidth(0, 15f);                                                             // Define column width
            wb.CurrentWorksheet.SetColumnWidth(1, 20f);                                                             // Define column width
            wb.CurrentWorksheet.SetColumnWidth(2, 35f);                                                             // Define column width

            wb.AddWorksheet("SplitYchars");                                                                         // Create new worksheet
            wb.CurrentWorksheet.SetHorizontalSplit(20f, new Address("C1"), Worksheet.WorksheetPane.bottomLeft);     // Split worksheet horizontally by characters
            wb.AddWorksheet("SplitYcols");                                                                          // Create new worksheet
            wb.CurrentWorksheet.SetHorizontalSplit(5, false, new Address("A6"), Worksheet.WorksheetPane.bottomLeft);// Split worksheet horizontally by rows
            wb.CurrentWorksheet.SetRowHeight(0, 10f);                                                               // Define row height
            wb.CurrentWorksheet.SetRowHeight(3, 35f);                                                               // Define row height
            wb.CurrentWorksheet.SetRowHeight(2, 25f);                                                               // Define row height

            wb.AddWorksheet("SplitXYchars");                                                                        // Create new worksheet
            wb.CurrentWorksheet.SetSplit(30f, 20f, new Address("D3"), Worksheet.WorksheetPane.bottomRight);         // Split worksheet horizontally and vertically by characters

            wb.AddWorksheet("SplitXYColRow");                                                                       // Create new worksheet
            wb.CurrentWorksheet.SetSplit(3, 10, false, new Address("D11"), Worksheet.WorksheetPane.bottomRight);    // Split worksheet horizontally and vertically by rows and columns

            wb.AddWorksheet("FreezeXcols");                                                                         // Create new worksheet
            wb.CurrentWorksheet.SetVerticalSplit(4, true, new Address("E1"), Worksheet.WorksheetPane.topRight);     // Split and freeze worksheet vertically by columns

            wb.AddWorksheet("FreezeYcols");                                                                         // Create new worksheet
            wb.CurrentWorksheet.SetHorizontalSplit(5, true, new Address("A6"), Worksheet.WorksheetPane.bottomLeft); // Split and freeze worksheet horizontally by rows

            wb.AddWorksheet("FreezeXYColRow");                                                                      // Create new worksheet
            wb.CurrentWorksheet.SetSplit(3, 10, true, new Address("D11"), Worksheet.WorksheetPane.bottomRight);     // Split and freeze worksheet horizontally and vertically by rows and columns

            wb.Save();                                                                                              // Save the workbook
        }

        private PaneSplitAndFreeze()
        {
            // Do not instantiate
        }
    }
}