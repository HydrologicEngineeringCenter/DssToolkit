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
        public bool HasExcelFile { get; set; }
        public bool HasDssFile { get; set; }
        public bool OverwriteSheets { get; set; }
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
                        DssPathList.Items.Add(path.FullPath);
                    HasDssFile = true;
                }
                DssFilePath.Text = openFileDialog.FileName;
            }
        }

        private void ExcelFileButton_Click(object sender, RoutedEventArgs e)
        {
            GetExcelFile();
        }

        private bool GetExcelFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx";
            if (openFileDialog.ShowDialog() == true)
            {
                ExcelReader er = new ExcelReader(openFileDialog.FileName);
                SheetList.Items.Clear();
                for (int i = 0; i < er.workbook.Worksheets.Count; i++)
                    SheetList.Items.Add(er.workbook.Worksheets[i].Name);
                HasExcelFile = true;
                ExcelFilePath.Text = openFileDialog.FileName;
                return true;
            }
            return false;
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
            if (CheckSelections())
            {
                List<string> sheets = GetImportExcelSheets();
                ExcelReader er = new ExcelReader(ExcelFilePath.Text);
                
                if (!File.Exists(DssFilePath.Text))
                {
                    SaveFileDialog browser = new SaveFileDialog();
                    browser.Title = "Select or Create DSS File";
                    browser.Filter = "DSS Files (*.dss)|*.dss";
                    if (browser.ShowDialog() != true)
                        return;
                    HasDssFile = true;
                    DssFilePath.Text = browser.FileName;
                }
                string filename = DssFilePath.Text;

                using (DssWriter w = new DssWriter(filename))
                {
                    
                    foreach (var sheet in sheets)
                    {
                        var t = er.CheckType(sheet);
                        if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                            w.Write(er.Read(sheet) as TimeSeries);
                        else if (t is RecordType.PairedData)
                            w.Write(er.Read(sheet) as PairedData);
                    }
                    RefreshDssPathList();

                    var result = MessageBox.Show(String.Format("DSS data has successfully been imported from {0} to {1}", er.workbook.FullName, w.Filename), 
                        "Import Successful", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                    if (result == MessageBoxResult.OK)
                    {
                        Process.Start("explorer.exe", Path.GetDirectoryName(filename));
                    }
                }
            }
        }

        private void RefreshDssPathList()
        {
            DssPathCollection c;
            using (DssReader r = new DssReader(DssFilePath.Text))
            {
                c = r.GetCatalog();
                DssPathList.Items.Clear();
                foreach (var path in c)
                    DssPathList.Items.Add(path.FullPath);
            }
        }

        private List<string> GetImportExcelSheets()
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
            if (CheckSelections())
            {
                List<string> sheets = GetExportExcelSheets();
                List<DssPath> paths = GetDssPaths();

                if (!File.Exists(ExcelFilePath.Text))
                {
                    SaveFileDialog browser = new SaveFileDialog();
                    browser.Title = "Select or Create Excel File";
                    browser.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx";
                    if (browser.ShowDialog() != true)
                        return;
                    HasExcelFile = true;
                    ExcelFilePath.Text = browser.FileName;
                }
                string filename = ExcelFilePath.Text;

                using (DssReader r = new DssReader(filename))
                {
                    object record;
                    ExcelWriter ew = new ExcelWriter(filename);
                    for (int i = 0; i < sheets.Count; i++)
                    {
                        DssPath p = new DssPath(paths[i].PathWithoutDate);
                        var type = r.GetRecordType(p);
                        if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                        {
                            record = r.GetTimeSeries(p);
                            ew.Write(record as TimeSeries, sheets[i]);
                        }
                        else if (type is RecordType.PairedData)
                        {
                            record = r.GetPairedData(p.FullPath);
                            ew.Write(record as PairedData, sheets[i]);
                        }
                    }
                    RefreshSheetList();

                    var result = MessageBox.Show(String.Format("DSS data has successfully been exported from {0} to {1}. Show excel file in File Explorer?", r.Filename, filename), 
                        "Export Successful", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                    if (result == MessageBoxResult.OK)
                    {
                        Process.Start("explorer.exe", Path.GetDirectoryName(filename));
                    }
                }
            }
        }

        private void RefreshSheetList()
        {
            ExcelReader er = new ExcelReader(ExcelFilePath.Text);
            SheetList.Items.Clear();
            for (int i = 0; i < er.workbook.Worksheets.Count; i++)
                SheetList.Items.Add(er.workbook.Worksheets[i].Name);
        }

        private List<DssPath> GetDssPaths()
        {
            List<DssPath> paths = new List<DssPath>();
            if (DssPathList.SelectedItems.Count != 0)
            {
                for (int i = 0; i < DssPathList.SelectedItems.Count; i++)
                {
                    paths.Add(new DssPath(DssPathList.SelectedItems[i].ToString()));
                }
            }
            else
            {
                for (int i = 0; i < DssPathList.Items.Count; i++)
                {
                    paths.Add(new DssPath(DssPathList.Items[i].ToString()));
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
            if (OverwriteSheets)
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

        private bool CheckSelections()
        {
            if (HasDssFile && HasExcelFile && DssPathList.SelectedItems.Count != SheetList.SelectedItems.Count)
            {
                MessageBox.Show("The amound of selected excel sheets and DSS paths do not match.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }

            return true;
        }

        private bool CanExport()
        {
            return HasDssFile;
        }

        private bool CanImport()
        {
            return HasExcelFile;
        }

        private void ManualImportButton_Click(object sender, RoutedEventArgs e)
        {
            if (ExcelFilePath.Text == "")
            {
                if (!GetExcelFile())
                    return;
            }
            ExcelReader r = new ExcelReader(ExcelFilePath.Text);
            DSSExcelManualImport s = new DSSExcelManualImport(r.workbook.FullName);
            s.ShowDialog();
        }

        private void ViewDssFileButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void ViewExcelFileButton_Click(object sender, RoutedEventArgs e)
        {

        }

        private void DssFilePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            ImportButton.IsEnabled = CanImport();
            ExportButton.IsEnabled = CanExport();
        }

        private void ExcelFilePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            ImportButton.IsEnabled = CanImport();
            ExportButton.IsEnabled = CanExport();
        }
    }
}
