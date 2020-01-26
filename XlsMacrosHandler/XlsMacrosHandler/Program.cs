using System.IO;

namespace XlsMacrosHandler
{
    static class Program
    {
        static void Main(string[] args)
        {
            WorkbookOperationsHelper.WriteMacro(Constants.OriginalExcelFilePath,
                                                Constants.ExcelWithMacrosFilePath,
                                                Constants.MacroSample);
            WorkbookOperationsHelper.ReadAllMacros(Constants.ExcelWithMacrosFilePath);
            WorkbookOperationsHelper.TestProcessCount(new[]
            {
                Constants.OriginalExcelFilePath,
                Constants.ExcelWithMacrosFilePath
            });
        }
    }
}
