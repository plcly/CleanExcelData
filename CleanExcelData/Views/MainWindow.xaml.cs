using ExcelHelper;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace CleanExcelData.Views
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var folderBase = txtFolder.Text;
            if (Directory.Exists(folderBase))
            {
                btnBegin.IsEnabled = false;
                var clearnFolder = System.IO.Path.Combine(folderBase, "Clearned");
                if (!Directory.Exists(clearnFolder))
                {
                    Directory.CreateDirectory(clearnFolder);
                }
                else
                {
                    var msgResult = MessageBox.Show("转换后将会覆盖Clearned文件夹里的文件", "注意", MessageBoxButton.YesNo);
                    if (msgResult != MessageBoxResult.Yes)
                    {
                        btnBegin.IsEnabled = true;
                        return;
                    }
                }
                var excelTool = new ExcelTool(folderBase, clearnFolder);
                Task.Run(() => excelTool.Clean(ExcuteMsg));
            }
            else
            {
                MessageBox.Show("文件目录不存在");
            }
        }
        public void ExcuteMsg(string msg, int percent)
        {
            Dispatcher.BeginInvoke((Action)delegate
           {
               if (msg == "完成")
               {
                   btnBegin.IsEnabled = true;
               }

               txtMsg.Text = msg;
               proBar.Value = percent;
           });

        }
    }
}
