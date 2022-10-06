using NanoXLSX.Demo.UseCases;
using System;

namespace NanoXLSX.Demo
{
    /// <summary>
    /// Main class to run all demos
    /// </summary>
    public class Program
    {

        static void Main(string[] args)
        {
            UseCases.BasicDemo.Run();
            UseCases.Read.Run();
            UseCases.Shortener.Run();
            UseCases.CellAndWorksheetSelection.Run();
            UseCases.Stream.Run();
#pragma warning disable CS4014
            UseCases.Async.Run(); // Normally, this method should be called with the await keyword (what is not possible here). Usually, async methods are called along the call stack with await until a terminal element (like a WPF button) is reached
#pragma warning restore CS4014
            UseCases.AddnextCell.Run();
            UseCases.DataTypes.Run();
            UseCases.CellDirectionsAndValues.Run();
            UseCases.ColumnWidthsRowHeights.Run();
            UseCases.StyleBasics.Run();
            UseCases.StyleAppending.Run();
            UseCases.ActiveAndSetStyle.Run();
            UseCases.CellRanges.Run();
            UseCases.Metadata.Run();
            UseCases.MergeCells.Run();
            UseCases.ProtectionAndPasswords.Run();
            UseCases.HidingRowsAndColumns.Run();
            UseCases.AutoFilter.Run();
            UseCases.SanitizingWorksheetNames.Run();
            UseCases.Formulas.Run();
            UseCases.PaneSplitAndFreeze.Run();
            UseCases.HidingWorkbooksAndWorksheets.Run();
        }

        private Program()
        {
            // Do not instantiate
        }
    }
}
