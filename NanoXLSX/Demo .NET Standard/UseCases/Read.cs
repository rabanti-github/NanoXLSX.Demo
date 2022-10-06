using System;
using System.Collections.Generic;
using System.IO;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    ///  This is a demo to read the previously created BasicDemo.xlsx file
    /// </summary>
    public class Read
    {

        public const string BASIC_DEMO_FILE = "BasicDemo.xlsx"; // It is assumed that the basic demo file, defined in the basicDemo class exists

        public static void Run()
        {
            Workbook wb = Workbook.Load(BASIC_DEMO_FILE);                                       // Load the workbook 'basic.xlsx' from file 
            Console.WriteLine("contains worksheet name: " + wb.CurrentWorksheet.SheetName);
            foreach (KeyValuePair<string, Cell> cell in wb.CurrentWorksheet.Cells)              // Cycle through cells of loaded workbook (first worksheet)
            {
                PrintCellInfo(cell.Value);                                                      // Show information about the loaded cell (helper function; not part of the API)
            }

            Address? lastAddress = wb.CurrentWorksheet.GetLastCellAddress();                     // Determine the last cell of the current worksheet
            Console.WriteLine("The last cell in the current worksheet is: " + lastAddress.ToString());
            Address? lastdataAddress = wb.CurrentWorksheet.GetLastDataCellAddress();             // Determine the last cell with data of the current worksheet
            Console.WriteLine("The last cell with data in the current worksheet is: " + lastdataAddress.ToString());

            // The same as stream
            using (FileStream fs = new FileStream(BASIC_DEMO_FILE, FileMode.Open))              // Open the 'basic.xlsx' as file stream  
            {
                Workbook wb2 = Workbook.Load(fs);                                               // Read the file stream
                Console.WriteLine("contains worksheet name: " + wb2.CurrentWorksheet.SheetName);
                foreach (KeyValuePair<string, Cell> cell in wb2.CurrentWorksheet.Cells)         // Cycle through cells of loaded workbook (first worksheet)
                {
                    PrintCellInfo(cell.Value);                                                  // Show information about the loaded cell (helper function; not part of the API)
                }
            }
        }

        /// <summary>
        /// Prints information about the passed cell
        /// </summary>
        /// <param name="cell">Cell to examine</param>
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

        private Read()
        {
            // Do not instantiate
        }

    }
}
