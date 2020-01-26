using System;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;

namespace XlsMacrosHandler
{
    public static class Constants
    {
        public static readonly string CurrentAppDirectory = Directory.GetCurrentDirectory();
        private const string InputFolderName = "input";
        private const string OutputFolderName = "output";
        private const string OriginalExcelFileName = "workbook_original.xlsx";
        private const string ExcelWithMacrosFileNameWithoutExtension = "workbook_WithMacroByCSharp";

        public static readonly string InputDirectory = Path.Combine(CurrentAppDirectory, InputFolderName);
        public static readonly string OutputDirectory = Path.Combine(CurrentAppDirectory, OutputFolderName);

        public static readonly string OriginalExcelFilePath = Path.Combine(InputDirectory, OriginalExcelFileName);
        public static readonly string ExcelWithMacrosFilePath = Path.Combine(OutputDirectory, ExcelWithMacrosFileNameWithoutExtension);

        private static readonly string MacroVbaSubName = "RemoveHeaderRows";

        public static readonly Macro MacroSample = new Macro
        {
            VbaSubName = MacroVbaSubName,
            Code = $@"Sub {MacroVbaSubName}()
'
' RemoveHeaderRows Macro
' Removes all header rows from the worksheet that don't have any value in column C
'

'

    On Error Resume Next
    Columns(""C"").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub"
        };
    }
}
