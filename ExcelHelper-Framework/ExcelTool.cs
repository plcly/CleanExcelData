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
        private IWorkbook _workbook;
        private IWorkbook _workbookNew;
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
                    try
                    {
                        ClearnFile(file, extension);
                    }
                    catch (Exception ex)
                    {
                        System.IO.File.AppendAllText("Log.txt", ex.Message + Environment.NewLine);
                    }
                }
            }

            //Task.Run(() =>
            //{
            //    Parallel.ForEach(files, file =>
            //     {
            //         var extension = Path.GetExtension(file.Name);
            //         if (extension == ".xls" || extension == ".xlsx")
            //         {
            //             try
            //             {
            //                 ClearnFile(file, extension);
            //             }
            //             catch (Exception ex)
            //             {
            //                 System.IO.File.AppendAllText("Log.txt", ex.Message + Environment.NewLine);
            //             }
            //         }
            //     });
            //}).Wait();

        }

        private void ClearnFile(FileInfo file, string extension)
        {
            if (extension == ".xls")
            {
                _workbook = new HSSFWorkbook(new NPOIFSFileSystem(file));
                _workbookNew = new HSSFWorkbook();
            }
            else
            {
                _workbook = new XSSFWorkbook(file);
                _workbookNew = new XSSFWorkbook();
            }
            for (int i = 0; i < _workbook.NumberOfSheets; i++)
            {
                ISheet sheet = _workbook.GetSheetAt(i);
                ISheet sheetNew = _workbookNew.CreateSheet(sheet.SheetName);

                for (int j = 0; j <= sheet.LastRowNum; j++)
                {
                    IRow row = sheet.GetRow(j);
                    if (row != null)
                    {
                        IRow rowNew = sheetNew.CreateRow(j);
                        for (int k = 0; k <= row.LastCellNum; k++)
                        {
                            ICell cell = row.GetCell(k);
                            if (cell != null)
                            {
                                ICell cellNew = rowNew.CreateCell(k);
                                try
                                {
                                    SetCellStyles(cell, cellNew);
                                    ClearnCell(cell, cellNew);
                                }
                                catch (Exception ex)
                                {
                                    File.AppendAllText("Log.txt",
                                        DateTime.Now + " : FileName:" + file.Name
                                        + "|SheetName:" + sheet.SheetName
                                        + "|Row:" + (j + 1)
                                        + "|Column:" + (k + 1) + Environment.NewLine
                                        + ex.Message + Environment.NewLine
                                        );
                                }

                            }
                        }
                    }
                }
            }
            var newFileName = Path.Combine(_clearnFolder, file.Name);
            if (File.Exists(newFileName))
            {
                File.Delete(newFileName);
            }
            using (FileStream fs = File.OpenWrite(newFileName)) //打开一个xls文件，如果没有则自行创建，如果存在myxls.xls文件则在创建是不要打开该文件！
            {
                _workbook.Write(fs);
            }
        }

        private void SetCellStyles(ICell cell, ICell cellNew)
        {
            cellNew.CellStyle.BorderLeft = cell.CellStyle.BorderLeft;
            cellNew.CellStyle.BorderDiagonal = cell.CellStyle.BorderDiagonal;
            cellNew.CellStyle.BorderDiagonalLineStyle = cell.CellStyle.BorderDiagonalLineStyle;
            cellNew.CellStyle.BorderDiagonalColor = cell.CellStyle.BorderDiagonalColor;
            cellNew.CellStyle.FillForegroundColor = cell.CellStyle.FillForegroundColor;
            cellNew.CellStyle.FillBackgroundColor = cell.CellStyle.FillBackgroundColor;
            cellNew.CellStyle.FillPattern = cell.CellStyle.FillPattern;
            cellNew.CellStyle.BottomBorderColor = cell.CellStyle.BottomBorderColor;
            cellNew.CellStyle.TopBorderColor = cell.CellStyle.TopBorderColor;
            cellNew.CellStyle.RightBorderColor = cell.CellStyle.RightBorderColor;
            cellNew.CellStyle.LeftBorderColor = cell.CellStyle.LeftBorderColor;
            cellNew.CellStyle.BorderBottom = cell.CellStyle.BorderBottom;
            cellNew.CellStyle.BorderTop = cell.CellStyle.BorderTop;
            cellNew.CellStyle.BorderRight = cell.CellStyle.BorderRight;
            cellNew.CellStyle.Rotation = cell.CellStyle.Rotation;
            cellNew.CellStyle.VerticalAlignment = cell.CellStyle.VerticalAlignment;
            cellNew.CellStyle.WrapText = cell.CellStyle.WrapText;
            cellNew.CellStyle.Alignment = cell.CellStyle.Alignment;
            cellNew.CellStyle.IsLocked = cell.CellStyle.IsLocked;
            cellNew.CellStyle.IsHidden = cell.CellStyle.IsHidden;
            cellNew.CellStyle.DataFormat = cell.CellStyle.DataFormat;
            cellNew.CellStyle.ShrinkToFit = cell.CellStyle.ShrinkToFit;
            cellNew.CellStyle.Indention = cell.CellStyle.Indention;

        }

        private void ClearnCell(ICell cell, ICell cellNew)
        {
            cellNew.SetCellType(cell.CellType);
            switch (cell.CellType)
            {
                case CellType.Unknown:
                case CellType.String:
                    if (cell.ToString().IndexOf('.') > -1)
                    {
                        if (double.TryParse(cell.ToString(), out double cellValue))
                        {
                            cellNew.SetCellType(CellType.Numeric);
                            cellNew.SetCellValue(cellValue);
                            return;
                        }
                    }
                    cellNew.SetCellValue(TrimCellValue(cell.ToString()));
                    break;
                case CellType.Numeric:                    
                    cellNew.SetCellValue(cell.NumericCellValue);
                    break;
                case CellType.Formula:                    
                    cellNew.SetCellFormula(cell.CellFormula);
                    break;
                case CellType.Boolean:                    
                    cellNew.SetCellValue(cell.BooleanCellValue);
                    break;
                case CellType.Error:                    
                    cellNew.SetCellErrorValue(cell.ErrorCellValue);
                    break;
                case CellType.Blank:
                default:
                    break;
            }
            
        }
        private string TrimCellValue(string cellvalue)
        {
            return cellvalue.Trim(new char[] { (char)160, (char)12288, (char)32 });
        }
    }
}
