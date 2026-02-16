using NanoXLSX;
using NanoXLSX.Extensions;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the basic usage of inline formatting with multiple text runs in a single cell
    /// </summary>
    public class InlineFormattingBasics
    {
        public static void Run()
        {
            Workbook workbook = new Workbook("InlineFormattingBasics.xlsx", "Sheet1");              // Create new workbook

            // Example 1: Multiple runs with plain text
            FormattedTextBuilder builder1 = new FormattedTextBuilder();
            builder1.AddRun("Hello");                                                               // Add first run
            builder1.AddRun(" ");                                                                   // Add space
            builder1.AddRun("World");                                                               // Add third run
            workbook.CurrentWorksheet.AddFormattedTextCell(builder1.Build(), 0, 0);                 // Add to cell A1

            // Example 2: Text with line breaks (using \n in the text)
            FormattedTextBuilder builder2 = new FormattedTextBuilder();
            builder2.AddRun("First line\nSecond line\nThird line");                                 // Line breaks with \n
            FormattedText formatted2 = builder2.Build();
            formatted2.WrapText = true;                                                             // Enable wrap text for line breaks to render
            workbook.CurrentWorksheet.AddFormattedTextCell(formatted2, 0, 1);                       // Add to cell A2

            // Example 3: Combining multiple text fragments
            FormattedTextBuilder builder3 = new FormattedTextBuilder();
            builder3.AddRun("Product: ");
            builder3.AddRun("PicoXLSX");
            builder3.AddRun("\n");                                                                  // Line break as separate run
            builder3.AddRun("Version: ");
            builder3.AddRun("4.0");
            FormattedText formatted3 = builder3.Build();
            formatted3.WrapText = true;                                                             // Enable wrap text
            workbook.CurrentWorksheet.AddFormattedTextCell(formatted3, 0, 2);                       // Add to cell A3

            workbook.Save();                                                                        // Save the workbook
        }

        private InlineFormattingBasics()
        {
            // Do not instantiate
        }
    }
}
