using Hec.Dss;
using Hec.Dss.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
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

namespace DSSExcel
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

        private void DssFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "DSS Files (*.dss)|*.dss";
            if (openFileDialog.ShowDialog() == true)
            {
                DssPathCollection c;
                using (DssReader r = new DssReader(openFileDialog.FileName))
                {
                    c = r.GetCatalog();
                    DssPathList.Items.Clear();
                    foreach (var path in c)
                    {
                        DssPathList.Items.Add(path.FullPath);
                    }
                }

                DssFilePath.Text = openFileDialog.FileName;
            }
        }

        private void ExcelFileButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                ExcelReader er = new ExcelReader(openFileDialog.FileName);
                SheetList.Items.Clear();
                foreach (var sheet in er.workbook.Worksheets)
                {
                    SheetList.Items.Add(sheet);
                }

                ExcelFilePath.Text = openFileDialog.FileName;
            }
        }
    }
}
