using NanoXLSX;
using System;
using System.IO;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This method shows how to save a workbook as stream
    /// </summary>
    public class Stream
    {
        public static void Run()
        {
            Workbook workbook = new Workbook(true);                                 // Create new workbook without file name
            workbook.CurrentWorksheet.AddNextCell("This is an example");            // Add cell A1
            workbook.CurrentWorksheet.AddNextCellFormula("=A1");                    // Add formula in cell B1
            workbook.CurrentWorksheet.AddNextCell(123456789);                       // Add cell C1

            using (MemoryStream ms = new MemoryStream())
            {
                workbook.SaveAsStream(ms, true);                                    // Save the workbook into the MemoryStream; IMPORTANT: Leave stream open (2nd parameter = true)
                ms.Position = 0;                                                    // Reset the stream position
                using (StreamReader sr = new StreamReader(ms))                      // Pass MemoryStream to StreamReader
                {
                    string binaryData = sr.ReadToEnd();                             // Write Stream to a string
                    Console.WriteLine("Number of symbols: " + binaryData.Length);   // Write some "useful" data
                }
            }
            using (FileStream fs = new FileStream("Stream.xlsx", FileMode.Create))  // Create a FileStream
            {
                workbook.SaveAsStream(fs);                                          // Save the workbook into the FileStream and close the stream after writing
            }
        }

        private Stream()
        {
            // Do not instantiate
        }
    }
}
