using NanoXLSX;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of meta data
    /// </summary>
    public class Metadata
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("Metadata.xlsx", "Sheet1");                                // Create new workbook

            workbook.WorkbookMetadata.Title = "PicoXLSX Metadata Demo";                                // Add meta data to workbook
            workbook.WorkbookMetadata.Subject = "This is a PicoXLSX metadata test";                     // Add meta data to workbook
            workbook.WorkbookMetadata.Creator = "PicoXLSX";                                             // Add meta data to workbook
            workbook.WorkbookMetadata.Keywords = "Keyword1;Keyword2;Keyword3";                          // Add meta data to workbook

            workbook.Save();                                                                            // Save the workbook
        }

        private Metadata()
        {
            // Do not instantiate
        }
    }
}
