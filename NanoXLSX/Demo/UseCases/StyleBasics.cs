using NanoXLSX.Styles;
using System;
using System.Collections.Generic;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the basic usage of several styles
    /// </summary>
    public class StyleBasics
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("StyleBasics.xlsx", "Sheet1");                                 // Create new workbook
            List<object> values = new List<object>() { "Header1", "Header2", "Header3" };                   // Create a List of values
            workbook.CurrentWorksheet.AddCellRange(values, new Address(0, 0), new Address(2, 0));           // Add a cell range to A4 - C4
            workbook.CurrentWorksheet.Cells["A1"].SetStyle(BasicStyles.Bold);                               // Assign predefined basic style to cell
            workbook.CurrentWorksheet.Cells["B1"].SetStyle(BasicStyles.Bold);                               // Assign predefined basic style to cell
            workbook.CurrentWorksheet.Cells["C1"].SetStyle(BasicStyles.Bold);                               // Assign predefined basic style to cell
            workbook.CurrentWorksheet.GoToNextRow();                                                        // Go to Row 2
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now);                                            // Add cell A2
            workbook.CurrentWorksheet.AddNextCell(2);                                                       // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(3);                                                       // Add cell B2
            workbook.CurrentWorksheet.GoToNextRow();                                                        // Go to Row 3
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now.AddDays(1));                                 // Add cell B1
            workbook.CurrentWorksheet.AddNextCell("B");                                                     // Add cell B2
            workbook.CurrentWorksheet.AddNextCell("C");                                                     // Add cell B3

            Style s = new Style();                                                                          // Create new style
            s.CurrentFill.SetColor("FF22FF11", Fill.FillType.FillColor);                                    // Set fill color
            s.CurrentFont.Underline = Font.UnderlineValue.Double;                                          // Set double underline
            s.CurrentCellXf.HorizontalAlign = CellXf.HorizontalAlignValue.Center;                           // Set alignment

            Style s2 = s.CopyStyle();                                                                       // Copy the previously defined style
            s2.CurrentFont.Italic = true;                                                                   // Change an attribute of the copied style

            workbook.CurrentWorksheet.Cells["B2"].SetStyle(s);                                              // Assign style to cell
            workbook.CurrentWorksheet.GoToNextRow();                                                        // Go to Row 3
            workbook.CurrentWorksheet.AddNextCell(DateTime.Now.AddDays(2));                                 // Add cell B1
            workbook.CurrentWorksheet.AddNextCell(true);                                                    // Add cell B2
            workbook.CurrentWorksheet.AddNextCell(false, s2);                                               // Add cell B3 with style in the same step 
            workbook.CurrentWorksheet.Cells["C2"].SetStyle(BasicStyles.BorderFrame);                        // Assign predefined basic style to cell

            Style s3 = BasicStyles.Strike;                                                                  // Create a style from a predefined style
            s3.CurrentCellXf.TextRotation = 45;                                                             // Set text rotation
            s3.CurrentCellXf.VerticalAlign = CellXf.VerticalAlignValue.Center;                              // Set alignment

            workbook.CurrentWorksheet.Cells["B4"].SetStyle(s3);                                             // Assign style to cell

            Style s4 = BasicStyles.BoldItalic;                                                              // Create a style from a predefined style
            s4.CurrentCellXf.HorizontalAlign = CellXf.HorizontalAlignValue.Right;                           // Set text alignment
            s4.CurrentCellXf.Indent = 4;                                                                    // Set indentation
            workbook.CurrentWorksheet.AddCell("Text", 1, 4, s4);                                            // Assign style to cell B5

            workbook.Save();                                                                                // Save the workbook
        }

        private StyleBasics()
        {
            // Do not instantiate
        }
    
    }
}