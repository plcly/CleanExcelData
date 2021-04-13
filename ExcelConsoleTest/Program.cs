using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
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
            Task.Run(() =>
            {
                Parallel.For(0, 20, i =>
                {
                    var row = sheet1.CreateRow(i);
                    var cell = row.CreateCell(i);
                    cell.SetCellValue("cell" + i);
              });
            }).Wait();

            //var row1 = sheet1.CreateRow(0);
            //var cell2 = row1.CreateCell(1);
            //cell2.SetCellValue("Test");
            //var row3 = sheet1.CreateRow(2);
            //var cell5 = row3.CreateCell(4);
            //cell5.SetCellValue(5555);
            if (File.Exists("Test.xlsx"))
            {
                File.Delete("Test.xlsx");
            }
            using (FileStream fs = File.OpenWrite("Test.xlsx")) 
            {
                wb.Write(fs);
            }
            Console.WriteLine("done");
            Console.Read();
        }
    }
}
