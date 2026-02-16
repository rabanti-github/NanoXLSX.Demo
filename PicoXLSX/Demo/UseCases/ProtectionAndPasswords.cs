using NanoXLSX;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of protecting cells, worksheet password protection and workbook protection
    /// </summary>
    public class ProtectionAndPasswords
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("ProtectionAndPasswords.xlsx", "Protected");                                   // Create new workbook
            workbook.CurrentWorksheet.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.Sort);               // Allow to sort sheet (worksheet is automatically set as protected)
            workbook.CurrentWorksheet.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.InsertRows);         // Allow to insert rows
            workbook.CurrentWorksheet.AddAllowedActionOnSheetProtection(Worksheet.SheetProtectionValue.SelectLockedCells);  // Allow to select cells (locked cells caused automatically to select unlocked cells)
            workbook.CurrentWorksheet.AddNextCell("Cell A1");                                                               // Add cell A1
            workbook.CurrentWorksheet.AddNextCell("Cell B1");                                                               // Add cell B1
            workbook.CurrentWorksheet.Cells["A1"].SetCellLockedState(false, true);                                          // Set the locking state of cell A1 (not locked but value is hidden when cell selected)
            workbook.AddWorksheet("PWD-Protected");                                                                         // Add a new worksheet
            workbook.CurrentWorksheet.AddCell("This worksheet is password protected. The password is:", 0, 0);              // Add cell A1
            workbook.CurrentWorksheet.AddCell("test123", 0, 1);                                                             // Add cell A2
            workbook.CurrentWorksheet.SetSheetProtectionPassword("test123");                                                // Set the password "test123"
            workbook.SetWorkbookProtection(true, true, true, null);                                                         // Set workbook protection (windows locked, structure locked, no password)
            workbook.Save();                                                                                                // Save the workbook
        }

        private ProtectionAndPasswords()
        {
            // Do not instantiate
        }
    }
}
