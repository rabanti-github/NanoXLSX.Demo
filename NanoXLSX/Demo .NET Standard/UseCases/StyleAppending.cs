using NanoXLSX.Styles;
using System;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This demo shows the usage of style appending
    /// </summary>
    public class StyleAppending
    {
        public static void Run()
        {
            Workbook wb = new Workbook("StyleAppending.xlsx", "styleAppending");                  // Create a new workbook

            Style style = new Style();                                                            // Create a new style
            style.Append(BasicStyles.Bold);                                                       // Append a basic style (bold) 
            style.Append(BasicStyles.Underline);                                                  // Append a basic style (underline) 
            style.Append(BasicStyles.Font("Arial Black", 20f));                                   // Append a basic style (custom font) 

            wb.WS.Value("THIS IS A TEST", style);                                                 // Add text and the appended style
            wb.WS.Down();                                                                         // Go to a new row

            Style chainedStyle = new Style()                                                      // Create a new style...
                .Append(BasicStyles.Underline)                                                    // ... and append another part (chaining underline)
                .Append(BasicStyles.ColorizedText("FF00FF"))                                      // ... and append another part (chaining colorized text)
                .Append(BasicStyles.ColorizedBackground("AAFFAA"));                               // ... and append another part (chaining colorized background)

            wb.WS.Value("Another test", chainedStyle);                                            // Add text and the appended style

            wb.Save();                                                                            // Save the workbook
        }

        private StyleAppending()
        {
            // Do not instantiate
        }
    }
}