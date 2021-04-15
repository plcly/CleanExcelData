using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CleanExcelData.ViewModels
{
    public class MainWindowViewModel: BindableBase
    {
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

        private bool _btnEnable;
        public bool BtnEnable
        {
            get { return _btnEnable; }
            set { SetProperty(ref _btnEnable, value); }
        }
        private string _btnContent= "开始转换";
        public string BtnContent
        {
            get { return _btnContent; }
            set { SetProperty(ref _btnContent, value); }
        }
        
    }
}
