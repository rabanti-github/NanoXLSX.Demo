using NanoXLSX.Styles;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This method shows the shortened style of writing cells
    /// </summary>
    public class Shortener
    {
        public static void Run()
        {
            Workbook wb = new Workbook("Shortener.xlsx", "Sheet1");     // Create a workbook (important: A worksheet must be created as well) 
            wb.WS.Value("Some Text");                                   // Add cell A1
            wb.WS.Value(58.55, BasicStyles.DoubleUnderline);            // Add a formatted value to cell B1
            wb.WS.Right(2);                                             // Move to cell E1   
            wb.WS.Value(true);                                          // Add cell E1
            wb.AddWorksheet("Sheet2");                                  // Add a new worksheet
            wb.CurrentWorksheet.CurrentCellDirection = Worksheet.CellDirection.RowToRow; // Change the cell direction
            wb.WS.Value("This is another text");                        // Add cell A1
            wb.WS.Formula("=A1");                                       // Add a formula in Cell A2
            wb.WS.Down();                                               // Go to cell A4
            wb.WS.Value("Formatted Text", BasicStyles.Bold);            // Add a formatted value to cell A4
            wb.Save();                                                  // Save the workbook
        }

        private Shortener()
        {
            // Do not instantiate
        }
    }
}
