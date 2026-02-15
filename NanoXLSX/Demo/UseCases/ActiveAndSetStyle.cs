using NanoXLSX.Styles;
using System.Collections.Generic;

namespace NanoXLSX.Demo.UseCases
{

    /// <summary>
    /// This demo shows the usage of active styles and the SetStyle function on a worksheet 
    /// </summary>
    public class ActiveAndSetStyle
    {
        internal static void Run()
        {
            Workbook workbook = new Workbook("ActiveStyle.xlsx", "Sheet1");                             // Create new workbook
            List<object> values = new List<object>() { "Header1", "Header2", "Header3" };               // Create a List of values
            workbook.CurrentWorksheet.SetActiveStyle(BasicStyles.BorderFrameHeader);                    // Assign predefined basic style as active style
            workbook.CurrentWorksheet.AddCellRange(values, "A1:C1");                                    // Add cell range

            values = new List<object>() { "Cell A2", "Cell B2", "Cell C2" };                            // Create a List of values
            workbook.CurrentWorksheet.SetActiveStyle(BasicStyles.BorderFrame);                          // Assign predefined basic style as active style
            workbook.CurrentWorksheet.AddCellRange(values, "A2:C2");                                    // Add cell range (using active style)

            values = new List<object>() { "Cell A3", "Cell B3", "Cell C3" };                            // Create a List of values
            workbook.CurrentWorksheet.ClearActiveStyle();                                               // Clear the active style 
            workbook.CurrentWorksheet.AddCellRange(values, "A3:C3");                                    // Add cell range (without style)

            Style style = new Style();                                                                  // Create a new style
            style.Append(BasicStyles.ColorizedBackground("FF0000"));                                    // Append a visible style component
            values = new List<object>() { "Cell A4", "Cell B4", "Cell C4" };                            // Create a List of values

            workbook.CurrentWorksheet.SetStyle("A4", style);                                            // Set style based on a string address
            workbook.CurrentWorksheet.SetStyle("B4:C4", style);                                         // Set style based on a string address range
            workbook.CurrentWorksheet.SetStyle(new Address(0, 7), style);                               // Set style based on a address object
            workbook.CurrentWorksheet.SetStyle(new Range(new Address("D1"), new Address(5, 8)), style); // Set style based on a range object (overwrites style on C3)
            workbook.CurrentWorksheet.SetStyle(new Address("G6"), new Address("H10"), style);           // Set style based on a two address objects as range

            workbook.Save();                                                                            // Save the workbook
        }

        private ActiveAndSetStyle()
        {
            // Do not instantiate
        }
    }
}