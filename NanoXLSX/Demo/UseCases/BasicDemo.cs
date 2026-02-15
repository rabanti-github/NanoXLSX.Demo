using System;


namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This is a demo to read the previously created basix.xlsx file
    /// </summary>
    public class BasicDemo
    {

        public static void Run()
        {
            Workbook workbook = new Workbook("BasicDemo.xlsx", "Sheet1");   // Create new workbook
            workbook.CurrentWorksheet.AddNextCell("Test");                  // Add cell A1
            workbook.CurrentWorksheet.AddNextCell(55.2);                    // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);            // Add cell C1
            workbook.Save();
        }

        private BasicDemo()
        {
            // Do not instantiate
        }

    }
}
