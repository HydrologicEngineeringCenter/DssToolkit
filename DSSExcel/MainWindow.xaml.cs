﻿using Hec.Dss;
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
                ImportButton.IsEnabled = CanImport();
                ExportButton.IsEnabled = CanExport();
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
                ImportButton.IsEnabled = CanImport();
                ExportButton.IsEnabled = CanExport();
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
                
                string path = DssFilePath.Text;
                if (!File.Exists(path))
                {
                    System.Windows.Forms.FolderBrowserDialog browser = new System.Windows.Forms.FolderBrowserDialog();
                    browser.ShowNewFolderButton = true;
                    browser.Description = "Select directory for new dss file.";
                    if (browser.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        return;
                    path = browser.SelectedPath + "\\" + "dss_excel" + ExcelTools.RandomString(10) + ".dss";
                
                }
                using (DssWriter w = new DssWriter(path))
                {
                    if (sheets.Count == 0)
                    {
                        for (int i = 0; i < er.Count; i++)
                        {
                            var t = er.CheckType(i);
                            if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                                w.Write(er.Read(i) as TimeSeries);
                            else if (t is RecordType.PairedData)
                                w.Write(er.Read(i) as PairedData);
                        }
                    }
                    else
                    {
                        foreach (var sheet in sheets)
                        {
                            var t = er.CheckType(sheet);
                            if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                                w.Write(er.Read(sheet) as TimeSeries);
                            else if (t is RecordType.PairedData)
                                w.Write(er.Read(sheet) as PairedData);
                        }
                    }
                    RefreshDssPathList();

                    var result = MessageBox.Show(String.Format("DSS data has successfully been imported from {0} to {1}", er.workbook.FullName, w.Filename), 
                        "Import Successful", MessageBoxButton.OKCancel, MessageBoxImage.Information);
                    if (result == MessageBoxResult.OK)
                    {
                        Process.Start("explorer.exe", Path.GetDirectoryName(path));
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

                string filename = ExcelFilePath.Text;
                if (!File.Exists(filename))
                {
                    System.Windows.Forms.FolderBrowserDialog browser = new System.Windows.Forms.FolderBrowserDialog();
                    browser.ShowNewFolderButton = true;
                    browser.Description = "Select directory for new excel file.";
                    if (browser.ShowDialog() != System.Windows.Forms.DialogResult.OK)
                        return;
                    filename = browser.SelectedPath + "\\" + "dss_excel" + ExcelTools.RandomString(10) + ".xlsx";
                }

                using (DssReader r = new DssReader(DssFilePath.Text))
                {
                    object record;
                    ExcelWriter ew = new ExcelWriter(filename);
                    if (sheets.Count == 0)
                    {
                        for (int i = 0; i < paths.Count; i++)
                        {
                            DssPath p = new DssPath(paths[i].PathWithoutDate);
                            p.Dpart = "";
                            var type = r.GetRecordType(p);
                            if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                            {
                                record = r.GetTimeSeries(p);
                                ew.Write(record as TimeSeries, i);
                            }
                            else if (type is RecordType.PairedData)
                            {
                                record = r.GetPairedData(p.FullPath);
                                ew.Write(record as PairedData, i);
                            }
                        }
                    }
                    else
                    {
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
                MessageBox.Show("The amound of selected excel sheets and DSS paths do not match.", "Error", MessageBoxButton.OK);
                return false;
            }

            return true;
        }

        private bool CanExport()
        {
            if (HasDssFile)
                return true;
            else
                return false;
        }

        private bool CanImport()
        {
            if (HasExcelFile)
                return true;
            else
                return false;
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
    }
}
