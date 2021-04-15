using ExcelHelper;
using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CleanExcelData.ViewModels
{
    public class MainWindowViewModel : BindableBase
    {
        #region Properties

        private string _title="Excel处理";
        public string Title
        {
            get { return _title; }
            set { SetProperty(ref _title, value); }
        }

        private string _folderPath;
        public string FolderPath
        {
            get { return _folderPath; }
            set { SetProperty(ref _folderPath, value); }
        }

        private int _progressBarValue;
        public int ProgressBarValue
        {
            get { return _progressBarValue; }
            set { SetProperty(ref _progressBarValue, value); }
        }

        private bool _btnEnable = true;
        public bool BtnEnable
        {
            get { return _btnEnable; }
            set { SetProperty(ref _btnEnable, value); }
        }
        private string _btnContent = "开始转换";
        public string BtnContent
        {
            get { return _btnContent; }
            set { SetProperty(ref _btnContent, value); }
        }

        private string _txtMsg;
        public string TxtMsg
        {
            get { return _txtMsg; }
            set { SetProperty(ref _txtMsg, value); }
        } 
        #endregion


        private DelegateCommand _cleanExcelCommand;
        public DelegateCommand CleanExcelCommand =>
            _cleanExcelCommand ?? (_cleanExcelCommand = new DelegateCommand(ExecuteCommandName));

        void ExecuteCommandName()
        {
            var folderBase = FolderPath;
            if (Directory.Exists(folderBase))
            {
                BtnEnable = false;
                var handledFolder = System.IO.Path.Combine(folderBase, "Handled");
                if (!Directory.Exists(handledFolder))
                {
                    Directory.CreateDirectory(handledFolder);
                }
                else
                {
                    var msgResult = MessageBox.Show("转换后将会覆盖Handled文件夹里的文件", "注意", MessageBoxButton.YesNo);
                    if (msgResult != MessageBoxResult.Yes)
                    {
                        BtnEnable = true;
                        return;
                    }
                }
                var excelTool = new ExcelTool(folderBase, handledFolder);
                Task.Run(() => excelTool.Handle(ExcuteMsg));
            }
            else
            {
                MessageBox.Show("文件目录不存在");
            }
        }
        public void ExcuteMsg(string msg, int percent)
        {
            if (msg == "完成")
            {
                BtnEnable = true;
            }

            TxtMsg = msg;
            ProgressBarValue = percent;
        }
    }
}
