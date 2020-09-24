using Hec.Dss;
using Hec.Dss.Excel;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel.Composition.Primitives;
using System.Diagnostics;
using System.IO;
using System.Resources;
using System.Windows;
using System.Windows.Controls;
using SpreadsheetGear;
using System.Windows.Media;

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public QuickImportVM GetDataContext
        {
            get { return (QuickImportVM)DataContext; }
        }
        public MainWindow()
        {
            InitializeComponent();
        }

        private void DssFileButton_Click(object sender, RoutedEventArgs e)
        {
            GetDssFile();
        }

        private void DataFileButton_Click(object sender, RoutedEventArgs e)
        {
            GetDataFile();
        }

        private bool GetDssFile()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "Select or Create DSS File";
            dialog.Filter = "DSS Files (*.dss)|*.dss";
            dialog.OverwritePrompt = false;
            if (dialog.ShowDialog() == true)
            {
                GetDataContext.DssFilePath = dialog.FileName;
                GetDataContext.GetAllPaths();
                return true;
            }
            return false;
        }

        private bool GetDataFile()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "Select or Create Excel File";
            dialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx|CSV Files (*.csv)|*.csv";
            dialog.OverwritePrompt = false;
            if (dialog.ShowDialog() == true)
            {
                GetDataContext.DataFilePath = dialog.FileName;
                GetDataContext.GetAllSheets();
                return true;
            }
            return false;
        }

        private void ImportButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckImportSelections())
            {
                GetDataContext.SelectedSheets = GetSelectedImportSheets();

                if (!File.Exists(GetDataContext.DssFilePath))
                    if (!GetDssFile()) { return; }

                GetDataContext.QuickImport();
                SheetList.SelectedItems.Clear();
                DssPathList.SelectedItems.Clear();
                var result = MessageBox.Show(String.Format("DSS data has successfully been imported from {0} to {1}. Show DSS file in File Explorer?",
                    GetDataContext.DataFilePath, GetDataContext.DssFilePath),
                    "Import Successful", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                if (result == MessageBoxResult.OK)
                    Process.Start("explorer.exe", Path.GetDirectoryName(GetDataContext.DssFilePath));
            }
        }

        private List<string> GetSelectedImportSheets()
        {
            var sheets = new List<string>();
            if (SheetList.SelectedItems.Count != 0)
            {
                for (int i = 0; i < SheetList.SelectedItems.Count; i++)
                    sheets.Add(SheetList.SelectedItems[i].ToString());
            }
            else
            {
                for (int i = 0; i < SheetList.Items.Count; i++)
                    sheets.Add(SheetList.Items[i].ToString());
            }
            return sheets;
        }

        private void ExportButton_Click(object sender, RoutedEventArgs e)
        {
            if (CheckExportSelections())
            {
                GetDataContext.SelectedSheets = GetExportExcelSheets();
                GetDataContext.SelectedPaths = GetSelectedDssPaths();

                if (!File.Exists(GetDataContext.DataFilePath))
                    if (!GetDataFile()) { return; }

                GetDataContext.QuickExport();
                SheetList.SelectedItems.Clear();
                DssPathList.SelectedItems.Clear();
                var result = MessageBox.Show(String.Format("DSS data has successfully been exported from {0} to {1}. Show data file in File Explorer?",
                    GetDataContext.DssFilePath, GetDataContext.DataFilePath),
                    "Export Successful", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                if (result == MessageBoxResult.OK)
                    Process.Start("explorer.exe", Path.GetDirectoryName(GetDataContext.DataFilePath));
            }
        }

        private List<string> GetSelectedDssPaths()
        {
            List<string> paths = new List<string>();
            if (DssPathList.SelectedItems.Count != 0)
            {
                for (int i = 0; i < DssPathList.SelectedItems.Count; i++)
                {
                    paths.Add(new DssPath(DssPathList.SelectedItems[i].ToString()).PathWithoutDate);
                }
            }
            else
            {
                for (int i = 0; i < DssPathList.Items.Count; i++)
                {
                    paths.Add(new DssPath(DssPathList.Items[i].ToString()).PathWithoutDate);
                }
            }
            return paths;
        }

        private List<string> GetExportExcelSheets()
        {
            if (SheetList.SelectedItems.Count != 0)
                return GetSelectedSheets();
            else
                return GenerateExportSheets();
        }

        private List<string> GenerateExportSheets()
        {
            var sheets = new List<string>();
            if (GetDataContext.OverwriteSheets)
            {
                int c = DssPathList.Items.Count > SheetList.Items.Count ? SheetList.Items.Count : DssPathList.Items.Count;
                for (int i = 0; i < c; i++)
                {
                    sheets.Add(SheetList.Items[i].ToString());
                }

                if (DssPathList.Items.Count > SheetList.Items.Count)
                {
                    for (int i = 0; i < Math.Abs(SheetList.Items.Count - DssPathList.Items.Count); i++)
                    {
                        sheets.Add("SheetImport" + ExcelTools.RandomString(3));
                    }
                }
            }
            else
            {
                for (int i = 0; i < DssPathList.Items.Count; i++)
                {
                    sheets.Add("SheetImport" + ExcelTools.RandomString(3));
                }
            }
            return sheets;
        }

        private List<string> GetSelectedSheets()
        {
            var sheets = new List<string>();
            for (int i = 0; i < SheetList.SelectedItems.Count; i++)
            {
                sheets.Add(SheetList.SelectedItems[i].ToString());
            }
            return sheets;
        }

        private bool CheckExportSelections()
        {
            if (SheetList.SelectedItems.Count != 0 && DssPathList.SelectedItems.Count != 0 &&
                SheetList.SelectedItems.Count != DssPathList.SelectedItems.Count)
            {
                MessageBox.Show("The amound of selected sheets and DSS paths do not match.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private bool CheckImportSelections()
        {
            if (DssPathList.SelectedItems.Count != 0 && SheetList.SelectedItems.Count != 0 && 
                DssPathList.SelectedItems.Count != SheetList.SelectedItems.Count)
            {
                MessageBox.Show("The amound of selected sheets and DSS paths do not match.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private void ManualImportButton_Click(object sender, RoutedEventArgs e)
        {
            if (!GetDataContext.HasDataFile)
            {
                if (!GetDataFile())
                    return;
            }
            ExcelReader er = new ExcelReader(GetDataContext.DataFilePath);
            DSSExcelManualImport s = new DSSExcelManualImport(er.workbook.FullName);
            s.ShowDialog();
        }

        private void ViewDssFileButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ViewDataFileButton_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
