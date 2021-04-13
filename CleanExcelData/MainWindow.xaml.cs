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

namespace CleanExcelData
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
            if (!string.IsNullOrWhiteSpace(txtFolder.Text))
            {
                var folderBase = txtFolder.Text;
                if (Directory.Exists(folderBase))
                {
                    var clearnFolder = System.IO.Path.Combine(folderBase, "Clearned");
                    if (!Directory.Exists(clearnFolder))
                    {
                        Directory.CreateDirectory(clearnFolder);
                    }
                    ExcelTool.Clean(folderBase, clearnFolder);
                }

            }
        }
    }
}
