using System;

namespace NanoXLSX.Demo.UseCases
{
    public class Metadata
    {
        /// <summary>
        /// This demo shows the usage of meta data 
        /// </summary>
        public static void Run()
        {
            Workbook workbook = new Workbook("Metadata.xlsx", "Sheet1");                                // Create new workbook

            workbook.WorkbookMetadata.Title = "Test 15th";                                              // Add meta data to workbook
            workbook.WorkbookMetadata.Subject = "This is the 15th NanoXLSX test";                       // Add meta data to workbook
            workbook.WorkbookMetadata.Creator = "NanoXLSX";                                             // Add meta data to workbook
            workbook.WorkbookMetadata.Keywords = "Keyword1;Keyword2;Keyword3";                          // Add meta data to workbook

            workbook.Save();                                                                            // Save the workbook
        }

        private Metadata()
        {
            // Do not instantiate
        }
    }
}