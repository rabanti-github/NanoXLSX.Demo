using System;
using System.Collections.Generic;

namespace NanoXLSX.Demo.UseCases
{

    /// <summary>
    /// This demo shows the usage of several data types, the method AddCell, more than one worksheet and the SaveAs method
    /// </summary>
    public class DataTypes
    {
        public static void Run()
        {
            Workbook workbook = new Workbook(false);                    // Create new workbook
            workbook.AddWorksheet("Sheet1");                            // Add a new Worksheet and set it as current sheet
            workbook.CurrentWorksheet.AddNextCell("月曜日");            // Add cell A1 (Unicode)
            workbook.CurrentWorksheet.AddNextCell(-987);                // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(false);               // Add cell C1
            workbook.CurrentWorksheet.GoToNextRow();                    // Go to Row 2
            workbook.CurrentWorksheet.AddNextCell(-123.456d);           // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(-123.789f);           // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);        // Add cell C3
            workbook.CurrentWorksheet.AddNextCell(new TimeSpan(23, 59, 59)); // Add cell D3
            workbook.AddWorksheet("Sheet2");                            // Add a new Worksheet and set it as current sheet
            workbook.CurrentWorksheet.AddCell("ABC", "A1");             // Add cell A1
            workbook.CurrentWorksheet.AddCell(779, 2, 1);               // Add cell C2 (zero based addresses: column 2=C, row 1=2)
            workbook.CurrentWorksheet.AddCell(false, 3, 2);             // Add cell D3 (zero based addresses: column 3=D, row 2=3)
            workbook.CurrentWorksheet.AddNextCell(0);                   // Add cell E3 (direction: column to column)
            List<object> values = new List<object>() { "V1", true, 16.8 }; // Create a List of values
            workbook.CurrentWorksheet.AddCellRange(values, "A4:C4");    // Add a cell range to A4 - C4
            workbook.SaveAs("DataTypes.xlsx");                          // Save the workbook
        }


        private DataTypes()
        {
            // Do not instantiate
        }

    }
}