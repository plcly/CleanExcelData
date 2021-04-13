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
            IWorkbook wb = new XSSFWorkbook();
            var sheet1 = wb.CreateSheet();
            var lastBlankRow = sheet1.CreateRow(21);
            var lastBlankCell = lastBlankRow.CreateCell(21, CellType.Blank);

            Stopwatch sw = new Stopwatch();
            sw.Start();
            //ForeachCreate(sheet1);//51516
            ParalleCreate(sheet1);//
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
            Console.WriteLine("done");
            Console.Read();
        }

        private static void ParalleCreate(ISheet sheet1)
        {
            var listCell = new HashSet<ICell>();
            for (int i = 0; i < 5000; i++)
            {
                var row = sheet1.CreateRow(i);
                for (int j = 0; j <5000; j++)
                {
                    var cell=row.CreateCell(j);
                    listCell.Add(cell);
                    
                }
            }
            Task.Run(() =>
            {
                Parallel.ForEach(listCell, cell =>
                {
                    cell.SetCellValue("TestCell");
                });
            }).Wait();
            
        }

        private static void ForeachCreate(ISheet sheet1)
        {
            for (int i = 0; i < 5000; i++)
            {
                var row = sheet1.CreateRow(i);
                for (int j = 0; j < 5000; j++)
                {
                    var cell = row.CreateCell(i);
                    cell.SetCellValue("TestCell");
                }
            }
            
        }
    }
}
