using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NanoXLSX.Demo.UseCases
{
    public class AddnextCell
    {
        /// <summary>
        /// This method shows the usage of AddNextCell with several data types and formulas. Furthermore, the several types of Addresses are demonstrated
        /// </summary>
        public static void Run()
        {
            Workbook workbook = new Workbook("AddnextCell.xlsx", "Sheet1");  // Create new workbook
            workbook.CurrentWorksheet.AddNextCell("Test");                   // Add cell A1
            workbook.CurrentWorksheet.AddNextCell(123);                      // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(true);                     // Add cell C1
            workbook.CurrentWorksheet.GoToNextRow();                         // Go to Row 2
            workbook.CurrentWorksheet.AddNextCell(123.456d);                 // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(123.789f);                 // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);             // Add cell C2
            workbook.CurrentWorksheet.AddNextCell(new TimeSpan(12, 50, 30)); // Add cell D2
            workbook.CurrentWorksheet.GoToNextRow();                         // Go to Row 3
            workbook.CurrentWorksheet.AddNextCellFormula("B1*22");           // Add cell A3 as formula (B1 times 22)
            workbook.CurrentWorksheet.AddNextCellFormula("ROUNDDOWN(A2,1)"); // Add cell B3 as formula (Floor A2 with one decimal place)
            workbook.CurrentWorksheet.AddNextCellFormula("PI()");            // Add cell C3 as formula (Pi = 3.14.... )
            workbook.AddWorksheet("Addresses");                                                 // Add new worksheet
            workbook.CurrentWorksheet.CurrentCellDirection = Worksheet.CellDirection.Disabled;  // Disable automatic addressing
            workbook.CurrentWorksheet.AddCell("Default", 0, 0);                                 // Add a value
            Address address = new Address(1, 0, Cell.AddressType.Default);                      // Create Address with default behavior
            workbook.CurrentWorksheet.AddCell(address.ToString(), 1, 0);                        // Add the string of the address
            workbook.CurrentWorksheet.AddCell("Fixed Column", 0, 1);                            // Add a value
            address = new Address(1, 1, Cell.AddressType.FixedColumn);                          // Create Address with fixed column
            workbook.CurrentWorksheet.AddCell(address.ToString(), 1, 1);                        // Add the string of the address
            workbook.CurrentWorksheet.AddCell("Fixed Row", 0, 2);                               // Add a value
            address = new Address(1, 2, Cell.AddressType.FixedRow);                             // Create Address with fixed row
            workbook.CurrentWorksheet.AddCell(address.ToString(), 1, 2);                        // Add the string of the address
            workbook.CurrentWorksheet.AddCell("Fixed Row and Column", 0, 3);                    // Add a value
            address = new Address(1, 3, Cell.AddressType.FixedRowAndColumn);                    // Create Address with fixed row and column
            workbook.CurrentWorksheet.AddCell(address.ToString(), 1, 3);                        // Add the string of the address
            workbook.Save();                                                                    // Save the workbook
        }

        private AddnextCell()
        {
            // Do not instantiate
        }
    }
}
