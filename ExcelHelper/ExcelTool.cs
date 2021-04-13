using System;
using System.IO;

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
            var files = Directory.GetFiles(_folderBase);
            foreach (var file in files)
            {
                TranslateFile(file);
            }
        }

        private void TranslateFile(string file)
        {
            //NPOI.XSSF.UserModel.XSSFWorkbook
        }
    }
}
