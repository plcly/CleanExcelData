using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public class ExcelTool
    {
        private string _folderBase;
        private string _handledFolder;
        public ExcelTool(string folderBase, string handledFolder)
        {
            _folderBase = folderBase;
            _handledFolder = handledFolder;
        }
        public void Handle(Action<string, int> excuteMsg)
        {
            var files = new DirectoryInfo(_folderBase).GetFiles();
            for (int i = 0; i < files.Length; i++)
            {
                var extension = Path.GetExtension(files[i].Name);
                if (extension == ".xls" || extension == ".xlsx")
                {
                    excuteMsg($"正在执行文件：{files[i].Name}({i}/{files.Length})" , i *100/ files.Length+5);
                    HandleFile(files[i], extension);
                }
            }
            excuteMsg($"完成", 100);
        }
        private void HandleFile(FileInfo file, string extension)
        {
            IWorkbook workbook;
            using (var fileStream = File.OpenRead(file.FullName))
            {
                if (extension == ".xls")
                {
                    workbook = new HSSFWorkbook(fileStream);
                }
                else
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                RunHandle(workbook);

                var newFileName = Path.Combine(_handledFolder, file.Name);
                if (File.Exists(newFileName))
                {
                    File.Delete(newFileName);
                }
                using (FileStream fs = File.OpenWrite(newFileName))
                {
                    workbook.Write(fs);
                }
                workbook.Close();
            }
        }

        private void RunHandle(IWorkbook workbook)
        {
            Task.Run(() =>
            {
                Parallel.For(0, workbook.NumberOfSheets, i =>
                {
                    ISheet sheet = workbook.GetSheetAt(i);
                    if (sheet != null)
                    {
                        Parallel.For(0, sheet.LastRowNum, j =>
                        {
                            IRow row = sheet.GetRow(j);
                            if (row != null)
                            {
                                Parallel.For(0, row.LastCellNum, k =>
                                {
                                    ICell cell = row.GetCell(k);
                                    if (cell != null)
                                    {
                                        CleanCell(cell);
                                    }
                                });
                            }
                        });
                    }
                });
            }).Wait();
        }

        private void CleanCell(ICell cell)
        {
            if (cell.CellType == CellType.String)
            {
                cell.SetCellValue(TrimCellValue(cell.ToString()));
            }
        }

        private string TrimCellValue(string cellvalue)
        {
            return cellvalue.Trim(new char[] { (char)160, (char)12288, (char)32 });
        }
    }
}
