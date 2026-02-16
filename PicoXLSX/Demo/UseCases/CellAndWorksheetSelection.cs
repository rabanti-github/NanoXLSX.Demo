using NanoXLSX;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of cell and worksheet selection
    /// </summary>
    public class CellAndWorksheetSelection
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("CellAndWorksheetSelection.xlsx", "sheet1");                // Create new workbook
            workbook.CurrentWorksheet.AddNextCell("Test");                                               // Add cell A1
            workbook.CurrentWorksheet.ClearSelectedCells();                                              // Clear any existing selections
            workbook.CurrentWorksheet.AddSelectedCells("A5:B10");                                        // Set the selection to the range A5:B10
            workbook.AddWorksheet("Sheet2");                                                             // Create new worksheet
            workbook.CurrentWorksheet.AddNextCell("Test2");                                              // Add cell A1
            Range range = new Range(new Address(1, 1), new Address(3, 3));                               // Create a cell range for the selection B2:D4
            workbook.CurrentWorksheet.ClearSelectedCells();                                              // Clear any existing selections
            workbook.CurrentWorksheet.AddSelectedCells(range);                                           // Set the selection to the range
            workbook.AddWorksheet("Sheet2", true);                                                       // Create new worksheet with already existing name; The name will be changed to Sheet21 due to auto-sanitizing (appending of 1)
            workbook.CurrentWorksheet.AddNextCell("Test3");                                              // Add cell A1
            workbook.CurrentWorksheet.ClearSelectedCells();                                              // Clear any existing selections
            workbook.CurrentWorksheet.AddSelectedCells(new Address(2, 2), new Address(4, 4));            // Set the selection to the range C3:E5
            workbook.SetSelectedWorksheet(1);                                                            // Set the second Tab as selected (zero-based: 1)
            workbook.Save();                                                                             // Save the workbook
        }

        private CellAndWorksheetSelection()
        {
            // Do not instantiate
        }
    }
}
