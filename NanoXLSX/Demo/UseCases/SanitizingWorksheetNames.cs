using System;

namespace NanoXLSX.Demo.UseCases
{
    public class SanitizingWorksheetNames
    {
        /// <summary>
        /// This demo shows the usage of worksheet name sanitizing and auto-sanitizing
        /// </summary>
        public static void Run()
        {
            Workbook workbook = new Workbook(false);                                                    // Create new workbook without worksheet
            string invalidSheetName = "Sheet?1";                                                        // ? is not allowed in the names of worksheets
            string sanitizedSheetName = Worksheet.SanitizeWorksheetName(invalidSheetName, workbook);    // Method to sanitize a worksheet name (replaces ? with _)
            workbook.AddWorksheet(sanitizedSheetName);                                                  // Add new worksheet
            workbook.SaveAs("SanitizingWorksheetNames.xlsx");                                           // Save the workbook

            Workbook workbook2 = new Workbook("SanitizingWorksheetNames.xlsx", "Sheet*1", true);  		// Create new workbook with invalid sheet name (*); Auto-Sanitizing will replace * with _
            workbook2.SaveAs("SanitizingWorksheetNames_b.xlsx");                                        // Save the workbook
        }

        private SanitizingWorksheetNames()
        {
            // Do not instantiate
        }
    }
}