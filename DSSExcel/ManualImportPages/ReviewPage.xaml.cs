using Hec.Dss;
using SpreadsheetGear;
using Hec.Dss.Excel;
using System.Windows;
using System.Windows.Controls;
using static Hec.Dss.Excel.ExcelTools;
using System.Collections.Generic;

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for DSSPathPage.xaml
    /// </summary>
    public partial class ReviewPage : UserControl
    {
        public UserControl PreviousPage;
        public RecordType currentRecordType;
        private bool ts_paths_generated = false;
        private bool pd_path_generated = false;
        public List<DssPath> ts_paths = new List<DssPath>();
        public DssPath pd_path = new DssPath();

        public string Apart 
        {
            get
            {
                return ApartTextBox.Text;
            }
        }
        public string Bpart 
        {
            get
            {
                return BpartTextBox.Text;
            }
        }
        public string Cpart 
        {
            get
            {
                return CpartTextBox.Text;
            }
        }
        public string Dpart 
        {
            get
            {
                return DpartTextBox.Text;
            }
        }
        public string Epart 
        {
            get
            {
                return EpartTextBox.Text;
            }
        }
        public string Fpart 
        {
            get
            {
                return FpartTextBox.Text;
            }
        }


        public string GetCurrentPath 
        { 
            get
            {
                return "/" + Apart +
                    "/" + Bpart +
                    "/" + Cpart +
                    "/" + Dpart +
                    "/" + Epart +
                    "/" + Fpart + "/";
            }
        }
        public ReviewPage()
        {
            InitializeComponent();
        }

        public event RoutedEventHandler ImportClick;
        public event RoutedEventHandler BackClick;
        private void DSSPathImportButton_Click(object sender, RoutedEventArgs e)
        {
            this.ImportClick?.Invoke(this, e);
        }

        private void DSSPathBackButton_Click(object sender, RoutedEventArgs e)
        {
            this.BackClick?.Invoke(this, e);
        }

        private void ShowTimeSeriesPaths(IRange values)
        {
            if (!ts_paths_generated)
                GenerateTimeSeriesPaths(values.ColumnCount);
            DataContext = ts_paths[0];
        }

        private void GenerateTimeSeriesPaths(int count)
        {
            ts_paths.Clear();
            for (int i = 0; i < count; i++)
            {
                DssPath path = new DssPath();
                path.Apart = "a" + RandomString(3);
                path.Bpart = "b" + RandomString(3);
                path.Cpart = "c" + RandomString(3);
                path.Dpart = "";
                path.Epart = "";
                path.Fpart = "TimeSeries" + RandomString(3);
                ts_paths.Add(path);
            }
            
            ts_paths_generated = true;
        }

        private void ShowPairedDataPath()
        {
            if (!pd_path_generated)
                GeneratePairedDataPath();
            DataContext = pd_path;
        }

        private void GeneratePairedDataPath()
        {
            pd_path.Apart = "a" + RandomString(3);
            pd_path.Bpart = "b" + RandomString(3);
            pd_path.Cpart = "c" + RandomString(3);
            pd_path.Dpart = "";
            pd_path.Epart = "e" + RandomString(3);
            pd_path.Fpart = "PairedData" + RandomString(3);
            pd_path_generated = true;
        }

        public void ShowPath(RecordType recordType, IRange range1, IRange range2)
        {
            IsReadOnly(false);
            if (recordType is RecordType.IrregularTimeSeries || recordType is RecordType.RegularTimeSeries)
            {
                currentRecordType = RecordType.RegularTimeSeries;
                ShowTimeSeriesPaths(range2);
            }
            else if (recordType is RecordType.PairedData)
            {
                currentRecordType = RecordType.PairedData;
                ShowPairedDataPath();
            }
            ShowRecordPreview(recordType, range1, range2);
            ExcelView.ActiveWorkbookSet.GetLock();
            ExcelView.ActiveWorksheet.Cells.Columns.AutoFit();
            ExcelView.ActiveWorkbookSet.ReleaseLock();
            IsReadOnly(true);
            InitializePathButtons();
        }

        private void IsReadOnly(bool option)
        {
            if (option)
            {
                ExcelView.ActiveWorkbookSet.GetLock();
                ExcelView.ActiveWorksheet.ProtectContents = true;
                ExcelView.ActiveWorkbookSet.ReleaseLock();
            }
            else
            {
                ExcelView.ActiveWorkbookSet.GetLock();
                ExcelView.ActiveWorksheet.ProtectContents = false;
                ExcelView.ActiveWorkbookSet.ReleaseLock();
            }
            
        }

        private void ShowRecordPreview(RecordType recordType, IRange range1, IRange range2)
        {
            if (recordType is RecordType.RegularTimeSeries || recordType is RecordType.IrregularTimeSeries)
                ShowTimeSeriesPreview(range1, range2);
            else if (recordType is RecordType.PairedData)
                ShowPairedDataPreview(range1, range2);
        }

        private void ShowTimeSeriesPreview(IRange dateTimes, IRange values)
        {
            ExcelView.ActiveWorkbookSet.GetLock();
            ExcelView.ActiveWorksheet.Cells.Clear();

            ExcelView.ActiveWorkbook.Worksheets[0].Cells[0, 0].Value = "Date/Time";
            for (int i = 1; i < dateTimes.RowCount + 1; i++)
            {
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[i, 0].Value = CellToString(dateTimes.Cells[i - 1, 0]);
            }

            for (int i = 1; i < values.ColumnCount + 1; i++)
            {
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[0, i].Value = "Value " + i.ToString();
                for (int j = 1; j < values.RowCount + 1; j++)
                {
                    ExcelView.ActiveWorkbook.Worksheets[0].Cells[j, i].Value = CellToString(values.Cells[j - 1, i - 1]);
                }
            }
            ExcelView.ActiveWorkbookSet.ReleaseLock();
        }

        private void ShowPairedDataPreview(IRange ordinates, IRange values)
        {
            ExcelView.ActiveWorkbookSet.GetLock();
            ExcelView.ActiveWorksheet.Cells.Clear();

            ExcelView.ActiveWorkbook.Worksheets[0].Cells[0, 0].Value = "Ordinates";
            for (int i = 1; i < ordinates.RowCount + 1; i++)
            {
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[i, 0].Value = CellToString(ordinates.Cells[i - 1, 0]);
            }

            for (int i = 1; i < values.ColumnCount + 1; i++)
            {
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[0, i].Value = "Value " + i.ToString();
                for (int j = 1; j < values.RowCount + 1; j++)
                {
                    ExcelView.ActiveWorkbook.Worksheets[0].Cells[j, i].Value = CellToString(values.Cells[j - 1, i - 1]);
                }
            }
            ExcelView.ActiveWorkbookSet.ReleaseLock();

        }

        public void ResetPaths()
        {
            ts_paths_generated = false;
            pd_path_generated = false;
            ts_paths.Clear();
            pd_path = new DssPath();
            PreviousPage = null;
        }

        private void ExcelView_ShowError(object sender, SpreadsheetGear.Windows.Controls.ShowErrorEventArgs e)
        {
            e.Handled = true;
        }

        private void prev_path_button_Click(object sender, RoutedEventArgs e)
        {
            if (currentRecordType is RecordType.RegularTimeSeries || currentRecordType is RecordType.IrregularTimeSeries)
            {
                int index = ts_paths.IndexOf(DataContext as DssPath);
                DataContext = ts_paths[--index];
                if (index == 0)
                    prev_path_button.IsEnabled = false;
                if (!next_path_button.IsEnabled)
                    next_path_button.IsEnabled = true;
            }
        }

        private void next_path_button_Click(object sender, RoutedEventArgs e)
        {
            if (currentRecordType is RecordType.RegularTimeSeries || currentRecordType is RecordType.IrregularTimeSeries)
            {
                int index = ts_paths.IndexOf(DataContext as DssPath);
                DataContext = ts_paths[++index];
                if (index == ts_paths.Count - 1)
                    next_path_button.IsEnabled = false;
                if (!prev_path_button.IsEnabled)
                    prev_path_button.IsEnabled = true;
            }
        }

        private void InitializePathButtons()
        {
            if (currentRecordType is RecordType.RegularTimeSeries || currentRecordType is RecordType.IrregularTimeSeries)
            {
                if (ts_paths.Count > 1)
                    next_path_button.IsEnabled = true;
            }
        }
    }
}
