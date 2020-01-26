using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using VBA = Microsoft.Vbe.Interop;

namespace XlsMacrosHandler
{
    public static class WorkbookOperationsHelper
    {
        public static void ReadAllMacros(string fileName)
        {
            var excel = new Excel.Application();
            var workbook = excel.Workbooks.Open(fileName, false, true, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Type.Missing);

            var project = workbook.VBProject;
            var projectName = project.Name;
            Console.WriteLine($"Project Name = {projectName}");

            foreach (var component in project.VBComponents)
            {
                VBA.VBComponent vbComponent = component as VBA.VBComponent;
                if (vbComponent != null)
                {
                    string componentName = vbComponent.Name;
                    Console.WriteLine("------------------------------------------");
                    Console.WriteLine($"vbComponentName = {componentName}");
                    Console.WriteLine("------------------------------------------");
                    var componentCode = vbComponent.CodeModule;
                    //Console.WriteLine($"vbComponentCodeModule = {componentCode}");
                    int componentCodeLines = componentCode.CountOfLines;
                    //string comments = string.Empty;
                    if (componentCodeLines > 0)
                    {
                        var code = componentCode.get_Lines(1, componentCodeLines);
                        Console.WriteLine($"code = {code}");

                        List<string> codeLines = code.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries).ToList();
                        var allMacros = new List<Macro>();
                        //string curMacroName = string.Empty, comments = String.Empty;
                        const string subStr = "Sub", endSubStr = "End Sub";
                        List<string> curMacroDescription = new List<string>(), curMacroCode = new List<string>();
                        Macro macro = null;
                        foreach (string codeLine in codeLines)
                        {
                            string trimmedCodeLine = codeLine.Trim();
                            if (trimmedCodeLine.StartsWith("'") && trimmedCodeLine.Length > 1)
                            {
                                curMacroDescription.Add(trimmedCodeLine.Substring(1));
                                curMacroCode.Add(codeLine);
                            }
                            else if (trimmedCodeLine.StartsWith(subStr) && trimmedCodeLine.Length > subStr.Length)
                            {
                                curMacroCode.Add(codeLine);
                                trimmedCodeLine = trimmedCodeLine.Substring(subStr.Length).Trim();
                                int openBracketIdx = trimmedCodeLine.IndexOf('(');
                                if (openBracketIdx > -1)
                                {
                                    macro = new Macro { Name = $"{componentName}.{trimmedCodeLine.Substring(0, openBracketIdx)}" };
                                    allMacros.Add(macro);
                                }
                            }
                            else if (trimmedCodeLine == endSubStr)
                            {
                                curMacroCode.Add(codeLine);
                                if (macro != null)
                                {
                                    macro.Description = string.Join(Environment.NewLine, curMacroDescription);
                                    macro.Code = string.Join(Environment.NewLine, curMacroCode);
                                    curMacroCode = new List<string>();
                                    curMacroDescription = new List<string>();
                                }
                            }
                            else
                            {
                                curMacroCode.Add(codeLine);
                            }
                        }
                        //comments = string.Join(Environment.NewLine, codeLines.Where(x => x.StartsWith("'") && x.Length > 1).Select(x => x.Substring(1).Trim()));

                        Console.WriteLine();
                        Console.WriteLine("------ RECAP --------------------------");
                        foreach (var curMacro in allMacros)
                        {
                            Console.WriteLine("---------------------------------------");
                            Console.WriteLine($"Macro name: {curMacro.Name}");
                            Console.WriteLine();
                            Console.WriteLine("Macro description:");
                            Console.WriteLine(string.Join(Environment.NewLine, curMacro.Description));
                            Console.WriteLine();
                            Console.WriteLine("Macro code:");
                            Console.WriteLine(string.Join(Environment.NewLine, curMacro.Code));
                        }
                    }
                    //Console.WriteLine($"Description = {comments}");

                    //int line = 1;
                    //while (line < componentCodeLines - 1)
                    //{
                    //    VBA.vbext_ProcKind procedureType;
                    //    string procedureName = componentCode.get_ProcOfLine(line, out procedureType);
                    //    if (!string.IsNullOrEmpty(procedureName))
                    //    {
                    //        string macroName = $"{componentName}.{procedureName}";
                    //        Console.WriteLine($"Macro Name = {macroName}");
                    //        int procedureLines = componentCode.get_ProcCountLines(procedureName, procedureType);
                    //        int procedureStartLine = componentCode.get_ProcStartLine(procedureName, procedureType);
                    //        int codeStartLine = componentCode.get_ProcBodyLine(procedureName, procedureType);
                    //        if (codeStartLine != procedureStartLine)
                    //        {
                    //            comments += componentCode.get_Lines(line, codeStartLine - procedureStartLine);
                    //        }

                    //        //int signatureLines = 1;
                    //        //while (componentCode.get_Lines(codeStartLine, signatureLines).EndsWith("_"))
                    //        //{
                    //        //    signatureLines++;
                    //        //}

                    //        //string signature = componentCode.get_Lines(codeStartLine, signatureLines);
                    //        //signature = signature.Replace("\n", string.Empty);
                    //        //signature = signature.Replace("\r", string.Empty);
                    //        //signature = signature.Replace("_", string.Empty);
                    //        line += procedureLines - 1;
                    //    }
                    //line++;
                    //}

                    //}
                    //}
                }
            }
            excel.Quit();
        }

        public static void TestProcessCount(string[] fileNames)
        {
            var exApp = new Excel.Application { Visible = false };
            var wbs = new List<Excel._Workbook>();
            try
            {
                foreach (var fileName in fileNames)
                {
                    var wb = exApp.Workbooks.Open(fileName, false, true, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Type.Missing);
                    wbs.Add(wb);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine($"Error in WriteMacro: {e.Message}, Stack Trace: {e.StackTrace}");
            }
            finally
            {
                foreach (var wb in wbs)
                {
                    wb.Close(false);
                }
                try
                {
                    exApp.Visible = false;
                    exApp.UserControl = false;
                }
                catch
                {
                    // ignored
                }

                // Gracefully exit out and destroy all COM objects to avoid hanging instances
                // of Excel.exe whether our method failed or not.
                exApp.Quit();

                foreach (var wb in wbs)
                {
                    if (wb != null) Marshal.ReleaseComObject(wb);
                }
                Marshal.ReleaseComObject(exApp);

                GC.Collect();
            }
        }

        public static void WriteMacro(string fileName,
            string fileNameToSave,
            Macro macro)
        {
            Excel.Application excel = null;
            Excel._Workbook workbook = null;
            VBA.VBComponent module = null;
            bool saveChanges = false;
            try
            {
                excel = new Excel.Application { Visible = false };
                workbook = excel.Workbooks.Open(fileName, false, true, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, false, true, Type.Missing);
                module = workbook.VBProject.VBComponents.Add(VBA.vbext_ComponentType.vbext_ct_StdModule);
                module.CodeModule.AddFromString(macro.Code);
                // Run the named VBA Sub that we just added.  In our sample, we named the Sub FormatSheet

                workbook.Application.Run(macro.VbaSubName, Missing.Value, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                                    Missing.Value, Missing.Value, Missing.Value);
                // Let loose control of the Excel instance

                excel.Visible = false;
                excel.UserControl = false;

                // Set a flag saying that all is well and it is ok to save our changes to a file.
                saveChanges = true;
                //  Save the file to disk
                //workbook.SaveAs(fileNameToSave, Excel.XlFileFormat.xlWorkbookNormal,
                //        null, null, false, false, Excel.XlSaveAsAccessMode.xlShared,
                //        false, false, null, null, null);

                workbook.SaveAs(fileNameToSave, Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled, Missing.Value,
    Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
    Excel.XlSaveConflictResolution.xlLocalSessionChanges, true,
    Missing.Value, Missing.Value, Missing.Value);

            }
            catch (Exception e)
            {
                Console.WriteLine($"Error in WriteMacro: {e.Message}, Stack Trace: {e.StackTrace}");
            }
            finally
            {
                try
                {
                    // Repeat excel.Visible and excel.UserControl releases just to be sure
                    // we didn't error out ahead of time.

                    if (excel != null)
                    {
                        excel.Visible = false;
                        excel.UserControl = false;
                    }
                    // Close the document and avoid user prompts to save if our
                    // method failed.
                    workbook?.Close(saveChanges, null, null);
                    excel?.Workbooks.Close();
                }
                catch
                {
                    // ignored
                }

                // Gracefully exit out and destroy all COM objects to avoid hanging instances
                // of Excel.exe whether our method failed or not.
                excel?.Quit();
                if (module != null) Marshal.ReleaseComObject(module);
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (excel != null) Marshal.ReleaseComObject(excel);   //This is used to kill the EXCEL.exe process

                GC.Collect();
            }
        }
    }
}
