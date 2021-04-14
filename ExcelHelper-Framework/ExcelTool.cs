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
        private string _clearnFolder;
        public ExcelTool(string folderBase, string clearnFolder)
        {
            _folderBase = folderBase;
            _clearnFolder = clearnFolder;
        }
        public void Clean()
        {
            var files = new DirectoryInfo(_folderBase).GetFiles();
            foreach (var file in files)
            {
                var extension = Path.GetExtension(file.Name);
                if (extension == ".xls" || extension == ".xlsx")
                {
                    ClearnFile(file, extension);
                }
            }
        }
        private void ClearnFile(FileInfo file, string extension)
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
                RunClearn(workbook);
                
                var newFileName = Path.Combine(_clearnFolder, file.Name);
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

        private void RunClearn(IWorkbook workbook)
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
