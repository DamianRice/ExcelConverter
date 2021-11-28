#region Info
// 2021
// 11
#endregion

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using Excel= Microsoft.Office.Interop.Excel;



namespace ExcelConverter
{
    public class ExcelHandler
    {
        private int process = 0;
        private Excel.Application _application = new Excel.Application();
        private Excel.Workbook _workbook;
        public event SetProcessHandler SetProcess;
        public event SetProcessHandler SetTotal;
        private Excel.Application Application
        {
            get => _application;
            set => _application = value;
        }

        private Excel.Workbook Workbook
        {
            get => _workbook;
            set => _workbook = value;
        }

        public bool ConvertExcel(string filePath, Excel.XlFileFormat saveFileFormat= Excel.XlFileFormat.xlOpenXMLWorkbook)
        {
            if (!File.Exists(filePath)) return false;

            string newFilePath = String.Empty;
            switch (saveFileFormat)
            {
                case Excel.XlFileFormat.xlOpenXMLStrictWorkbook:
                case Excel.XlFileFormat.xlOpenXMLWorkbook:
                    newFilePath = Path.ChangeExtension(filePath, @"xlsx");
                    break;
                case Excel.XlFileFormat.xlExcel8:
                case Excel.XlFileFormat.xlExcel9795:
                case Excel.XlFileFormat.xlExcel7| Excel.XlFileFormat.xlExcel5:
                    newFilePath = Path.ChangeExtension(filePath, @"xls");
                    break;
            }

            //新文件路径不为空且不存在同名文件
            if (!string.IsNullOrEmpty(newFilePath)&& !File.Exists(newFilePath))
            {
                Workbook = Application.Workbooks.Open(filePath);
                // todo 重名文件判断
                Workbook.SaveAs(
                    Filename: newFilePath,
                    FileFormat: saveFileFormat,
                    Password: Missing.Value,
                    WriteResPassword: Missing.Value,
                    ReadOnlyRecommended: Missing.Value,
                    CreateBackup: Missing.Value,
                    AccessMode: Excel.XlSaveAsAccessMode.xlNoChange,
                    ConflictResolution: Excel.XlSaveConflictResolution.xlUserResolution,
                    AddToMru: Missing.Value,
                    TextCodepage: Missing.Value,
                    TextVisualLayout: Missing.Value,
                    Local: Missing.Value
                );
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook);
                Workbook = null;
                GC.Collect();
                return true;
            }

            return false;

        }

        public Dictionary<string,bool> ConvertDirection(string dirPath, Excel.XlFileFormat saveFileFormat = Excel.XlFileFormat.xlOpenXMLWorkbook)
        {
            if (!Directory.Exists(dirPath)) return new Dictionary<string, bool> {{dirPath,false}};

            string filter = String.Empty;
            Dictionary<string, bool> result = new Dictionary<string, bool>();

            switch (saveFileFormat)
            {
                case Excel.XlFileFormat.xlOpenXMLStrictWorkbook:
                case Excel.XlFileFormat.xlOpenXMLWorkbook:
                    filter = @"^.+\.(xls)$";
                    break;
                case Excel.XlFileFormat.xlExcel8:
                case Excel.XlFileFormat.xlExcel9795:
                case Excel.XlFileFormat.xlExcel7 | Excel.XlFileFormat.xlExcel5:
                    filter = @"^.+\.(xlsx)$";
                    break;
            }

            if (!string.IsNullOrEmpty(filter))
            {
                var resultEntries = Directory.GetFiles(dirPath).Where(file => Regex.IsMatch(file, filter)).ToArray();
                SetTotal(resultEntries.Length);
                foreach (string resultEntry in resultEntries)
                {
                    process++;
                    var res = ConvertExcel(resultEntry, saveFileFormat);
                    result.Add(resultEntry, res);
                    SetProcess(process);
                }

                return result;
            }
            
            return new Dictionary<string, bool> { { dirPath, false } }; ;
        }

        public void CleanGC()
        {
            try
            {
                Workbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbook);
            }
            
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                Application.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(Application);
                Application = null;
                Workbook = null;
                GC.Collect();
            }

        }
    }
}