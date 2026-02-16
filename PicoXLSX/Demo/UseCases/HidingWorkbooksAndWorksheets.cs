using NanoXLSX;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of hiding workbooks and worksheets
    /// </summary>
    public class HidingWorkbooksAndWorksheets
    {
        public static void Run()
        {
            Workbook wb = new Workbook("HidingWorkbooksAndWorksheets.xlsx", "hiddenWB");    // Create a new workbook
            wb.CurrentWorksheet.AddNextCell("Hidden Workbook");
            wb.Hidden = true;                                                               // Set the workbook hidden (Set visible again in another, visible workbook)
            wb.Save();                                                                      // Save the workbook

            Workbook wb2 = new Workbook("HiddenWorksheet.xlsx", "visible");                 // Create a new workbook
            wb2.CurrentWorksheet.AddNextCell("Visible Worksheet");
            wb2.AddWorksheet("hidden");                                                     // Create new worksheet
            wb2.CurrentWorksheet.AddNextCell("Hidden Worksheet");
            wb2.CurrentWorksheet.Hidden = true;                                             // Set the current worksheet hidden
            wb2.Save();                                                                     // Save the workbook
        }

        private HidingWorkbooksAndWorksheets()
        {
            // Do not instantiate
        }
    }
}
