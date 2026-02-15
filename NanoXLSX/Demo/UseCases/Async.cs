using System.Threading.Tasks;

namespace NanoXLSX.Demo.UseCases
{
    /// <summary>
    /// This method shows how to save a workbook asynchronous
    /// </summary>
    public class Async
    {
        public static async Task Run()
        {
            Workbook workbook = new Workbook("Async.xlsx", "shet1");        // Create new workbook with file name
            workbook.WS.Value("Some text");                                 // Add cell A1
            workbook.WS.Value(222);                                         // Add cell B1
            workbook.WS.Formula("=A2");                                     // Add cell C1
            await workbook.SaveAsync();                                     // Save async
        }

        private Async()
        {
            // Do not instantiate
        }
    }
}
