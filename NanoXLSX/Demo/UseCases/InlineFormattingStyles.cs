using NanoXLSX.Extensions;
using NanoXLSX.Styles;
using System;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows inline formatting with styled text runs (bold, italic, colors, sizes, etc.)
    /// </summary>
    public class InlineFormattingStyles
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("InlineFormattingStyles.xlsx", "Sheet1");              // Create new workbook

            // Example 1: Bold and italic text using Font objects
            Font boldFont = new Font { Bold = true };
            Font italicFont = new Font { Italic = true };

            FormattedTextBuilder builder1 = new FormattedTextBuilder();
            builder1.AddRun("Bold text", boldFont);
            builder1.AddRun(" and ");
            builder1.AddRun("italic text", italicFont);
            workbook.CurrentWorksheet.AddFormattedTextCell(builder1.Build(), 0, 0);                 // Add to cell A1

            // Example 2: Different colors and sizes
            Font redFont = new Font { ColorValue = "FFFF0000", Size = 14 };
            Font blueFont = new Font { ColorValue = "FF0000FF", Size = 18 };

            FormattedTextBuilder builder2 = new FormattedTextBuilder();
            builder2.AddRun("Red ", redFont);
            builder2.AddRun("and ", new Font());
            builder2.AddRun("Blue", blueFont);
            workbook.CurrentWorksheet.AddFormattedTextCell(builder2.Build(), 0, 1);                 // Add to cell A2

            // Example 3: Using InlineStyleBuilder for fluent API
            FormattedTextBuilder builder3 = new FormattedTextBuilder();
            builder3.AddRun("Bold+Italic", sb => sb.Bold().Italic());
            builder3.AddRun(" ");
            builder3.AddRun("Large Green", sb => sb.Size(16).Color("FF00FF00"));
            builder3.AddRun(" ");
            builder3.AddRun("Underline", sb => sb.Underline());
            workbook.CurrentWorksheet.AddFormattedTextCell(builder3.Build(), 0, 2);                 // Add to cell A3

            // Example 4: Complex formatting with multiple attributes
            FormattedTextBuilder builder4 = new FormattedTextBuilder();
            builder4.AddRun("IMPORTANT: ", sb => sb.Bold().Size(14).Color("FFFF0000"));
            builder4.AddRun("This is a ");
            builder4.AddRun("formatted", sb => sb.Italic().Underline());
            builder4.AddRun(" message.");
            workbook.CurrentWorksheet.AddFormattedTextCell(builder4.Build(), 0, 3);                 // Add to cell A4

            workbook.Save();                                                                        // Save the workbook
        }

        private InlineFormattingStyles()
        {
            // Do not instantiate
        }
    }
}
