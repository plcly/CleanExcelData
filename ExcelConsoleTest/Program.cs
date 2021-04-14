using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelConsoleTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Begin");
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //ReadAndUpdate();
            ReadAndUpdateFor();
            sw.Stop();
            Console.WriteLine(sw.ElapsedMilliseconds);
            Console.WriteLine("done");
            Console.Read();
        }
        static void ReadAndUpdate()
        {
            using (var fileStream = File.OpenRead(Path.Combine(Directory.GetCurrentDirectory(), "Test.xlsx")))
            {
                IWorkbook _workbook = new XSSFWorkbook(fileStream);
                Task.Run(() =>
                {
                    Parallel.For(0, _workbook.NumberOfSheets, i =>
                    {
                        ISheet sheet = _workbook.GetSheetAt(i);
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
                                            if (cell.CellType == CellType.String)
                                            {
                                                cell.SetCellValue("NewCell");
                                            }
                                        }
                                    });
                                }
                            });
                        }
                    });
                }).Wait();
                var newFileName = Path.Combine(Directory.GetCurrentDirectory(), "TestNew.xlsx");
                if (File.Exists(newFileName))
                {
                    File.Delete(newFileName);
                }
                using (FileStream fs = File.OpenWrite(newFileName))
                {
                    _workbook.Write(fs);
                }
            }
        }
        static void ReadAndUpdateFor()
        {
            using (var fileStream = File.OpenRead(Path.Combine(Directory.GetCurrentDirectory(), "Test.xlsx")))
            {
                IWorkbook _workbook = new XSSFWorkbook(fileStream);

                Task.Run(() =>
                {
                    for (int i = 0; i < _workbook.NumberOfSheets; i++)
                    {
                        ISheet sheet = _workbook.GetSheetAt(i);
                        if (sheet != null)
                        {
                            for (int j = 0; j < sheet.LastRowNum; j++)
                            {
                                IRow row = sheet.GetRow(j);
                                if (row != null)
                                {
                                    for (int k = 0; k < row.LastCellNum; k++)
                                    {
                                        ICell cell = row.GetCell(k);
                                        if (cell != null)
                                        {
                                            if (cell.CellType == CellType.String)
                                            {
                                                cell.SetCellValue("NewCell");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    };
                }).Wait();
                var newFileName = Path.Combine(Directory.GetCurrentDirectory(), "TestNew.xlsx");
                if (File.Exists(newFileName))
                {
                    File.Delete(newFileName);
                }
                using (FileStream fs = File.OpenWrite(newFileName))
                {
                    _workbook.Write(fs);
                }
                _workbook.Close();
            }

        }
        static void CreateExcel()
        {
            IWorkbook wb = new XSSFWorkbook();
            var sheet1 = wb.CreateSheet();
            int rowNum = 5000;
            int columnNum = 100;
            Stopwatch sw = new Stopwatch();
            sw.Start();
            //ForeachCreate(sheet1, rowNum, columnNum);//27180
            //TaskCreate(sheet1, rowNum, columnNum);//24484
            //TaskCreateAddList(sheet1, rowNum, columnNum);//24883
            ParalleCreate(sheet1, rowNum, columnNum);//24226
            sw.Stop();
            Console.WriteLine(sw.ElapsedMilliseconds);
            //var row1 = sheet1.CreateRow(0);
            //var cell2 = row1.CreateCell(1);
            //cell2.SetCellValue("Test");
            //var row3 = sheet1.CreateRow(2);
            //var cell5 = row3.CreateCell(4);
            //cell5.SetCellValue(5555);
            var fileName = Path.Combine(Directory.GetCurrentDirectory(), "Test.xlsx");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
            using (FileStream fs = File.OpenWrite(fileName))
            {
                wb.Write(fs);
            }
        }
        private static void TaskCreateAddList(ISheet sheet1, int rowNum, int columnNum)
        {
            List<ICell> listCell = new List<ICell>();
            Task.Run(() =>
            {
                for (int i = 0; i < rowNum; i++)
                {
                    var row = sheet1.CreateRow(i);
                    for (int j = 0; j < columnNum; j++)
                    {
                        var cell = row.CreateCell(j);
                        listCell.Add(cell);
                        cell.SetCellValue("TestCell");
                    }
                }
            }).Wait();
        }

        private static void TaskCreate(ISheet sheet1, int rowNum, int columnNum)
        {
            Task.Run(() =>
            {
                for (int i = 0; i < rowNum; i++)
                {
                    var row = sheet1.CreateRow(i);
                    for (int j = 0; j < columnNum; j++)
                    {
                        var cell = row.CreateCell(j);
                        cell.SetCellValue("TestCell");
                    }
                }
            }).Wait();
        }

        private static void ParalleCreate(ISheet sheet1, int rowNum, int columnNum)
        {
            Task.Run(() =>
            {
                for (int i = 0; i < rowNum; i++)
                {
                    var row = sheet1.CreateRow(i);
                    Parallel.For(0, columnNum, j =>
                     {
                         var cell = row.CreateCell(j);
                         cell.SetCellValue("TestCell");
                     });
                }
            }).Wait();

        }

        private static void ForeachCreate(ISheet sheet1, int rowNum, int columnNum)
        {
            for (int i = 0; i < rowNum; i++)
            {
                var row = sheet1.CreateRow(i);
                for (int j = 0; j < columnNum; j++)
                {
                    var cell = row.CreateCell(j);
                    cell.SetCellValue("TestCell");
                }
            }

        }
    }
}
