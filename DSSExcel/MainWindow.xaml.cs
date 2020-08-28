using Hec.Dss;
using Hec.Dss.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition.Primitives;
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

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public bool HasExcelFile { get; set; }
        public bool HasDssFile { get; set; }
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
                    HasDssFile = true;
                }

                DssFilePath.Text = openFileDialog.FileName;
                ImportButton.IsEnabled = CanImport();
                ExportButton.IsEnabled = CanExport();
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
                HasExcelFile = true;
                ExcelFilePath.Text = openFileDialog.FileName;
                ImportButton.IsEnabled = CanImport();
                ExportButton.IsEnabled = CanExport();
            }
        }

        private void SheetList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ImportButton.IsEnabled = CanImport();
        }

        private void DssPathList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ExportButton.IsEnabled = CanExport();
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private bool CanImport()
        {
            if (HasDssFile && SheetList.SelectedItems.Count > 0)
                return true;
            else
                return false;
        }

        private bool CanExport()
        {
            if (HasExcelFile && DssPathList.SelectedItems.Count > 0)
                return true;
            else
                return false;
        }
    }
}
