using NanoXLSX;
using System;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This is a very basic demo (adding three values and save the workbook)
    /// </summary>
    public class BasicDemo
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("BasicDemo.xlsx", "Sheet1");   // Create new workbook
            workbook.CurrentWorksheet.AddNextCell("Test");                  // Add cell A1
            workbook.CurrentWorksheet.AddNextCell(55.2);                    // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);            // Add cell C1
            workbook.Save();                                               // Save the workbook
        }

        private BasicDemo()
        {
            // Do not instantiate
        }
    }
}
