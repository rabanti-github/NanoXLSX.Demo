using NanoXLSX;
using System.Threading.Tasks;

namespace PicoXLSX.Demo.UseCases
{
    /// <summary>
    /// This use case shows how to save a workbook asynchronously
    /// </summary>
    public class SaveAsync
    {
        public static async Task Run()
        {
            Workbook workbook = new Workbook("SaveAsync.xlsx", "sheet1");       // Create new workbook with file name
            workbook.WS.Value("Some text");                                     // Add cell A1
            workbook.WS.Value(222);                                             // Add cell B1
            workbook.WS.Formula("=A2");                                         // Add cell C1
            await workbook.SaveAsync();                                         // Save async
        }

        private SaveAsync()
        {
            // Do not instantiate
        }
    }
}
