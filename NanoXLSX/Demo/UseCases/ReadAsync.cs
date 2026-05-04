using NanoXLSX.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This use case shows how to load a workbook asynchronously using WorkbookReader.LoadAsync
    /// </summary>
    public class ReadAsync
    {
        public const string BASIC_DEMO_FILE = "BasicDemo.xlsx"; // It is assumed that the basic demo file, defined in the BasicDemo class exists

        public static async Task Run()
        {
            Console.WriteLine("Loading workbook asynchronously from file...");
            Workbook wb = await WorkbookReader.LoadAsync(BASIC_DEMO_FILE);                          // Load the workbook asynchronously from file
            Console.WriteLine("Contains worksheet name: " + wb.CurrentWorksheet.SheetName);
            foreach (KeyValuePair<string, Cell> cell in wb.CurrentWorksheet.Cells)                 // Cycle through cells of loaded workbook (first worksheet)
            {
                PrintCellInfo(cell.Value);
            }

            Address? lastAddress = wb.CurrentWorksheet.GetLastCellAddress();                        // Determine the last cell of the current worksheet
            Console.WriteLine("The last cell in the current worksheet is: " + lastAddress.ToString());
            Address? lastDataAddress = wb.CurrentWorksheet.GetLastDataCellAddress();                // Determine the last cell with data of the current worksheet
            Console.WriteLine("The last cell with data in the current worksheet is: " + lastDataAddress.ToString());

            Console.WriteLine("\nLoading workbook asynchronously from stream...");
            using (FileStream fs = new FileStream(BASIC_DEMO_FILE, FileMode.Open))                 // Open the file as a stream
            {
                Workbook wb2 = await WorkbookReader.LoadAsync(fs);                                 // Load the workbook asynchronously from stream
                Console.WriteLine("Contains worksheet name: " + wb2.CurrentWorksheet.SheetName);
                foreach (KeyValuePair<string, Cell> cell in wb2.CurrentWorksheet.Cells)            // Cycle through cells of loaded workbook (first worksheet)
                {
                    PrintCellInfo(cell.Value);
                }
            }
        }

        private static void PrintCellInfo(Cell cell)
        {
            Console.Write("Cell address: " + cell.CellAddress);
            Console.Write(" - content: '" + cell.Value + "'");
            Console.Write(" (Type of: " + cell.DataType);
            Console.Write(" / Has style: ");
            if (cell.CellStyle == null)
            {
                Console.Write("none)\n");
            }
            else
            {
                Console.Write(cell.CellStyle.CurrentNumberFormat.Number + ")\n");
            }
        }

        private ReadAsync()
        {
            // Do not instantiate
        }
    }
}
