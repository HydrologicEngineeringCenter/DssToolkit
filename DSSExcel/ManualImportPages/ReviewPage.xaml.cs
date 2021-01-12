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
        public List<DssPath> ts_paths = new List<DssPath>();
        public DssPath pd_path = new DssPath();
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
            GenerateTimeSeriesPaths(values.ColumnCount);
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
                path.Fpart = "TimeSeries";
                ts_paths.Add(path);
            }
        }

        private void ShowPairedDataPath()
        {
            GeneratePairedDataPath();
        }

        private void GeneratePairedDataPath()
        {
            pd_path = new DssPath();
            pd_path.Apart = "a" + RandomString(3);
            pd_path.Bpart = "b" + RandomString(3);
            pd_path.Cpart = "c" + RandomString(3);
            pd_path.Dpart = "";
            pd_path.Epart = "e" + RandomString(3);
            pd_path.Fpart = "PairedData";
        }

        public void SetupReviewPage(RecordType recordType, IRange range1, IRange range2)
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

            int rowStart = 0;
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "A";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "B";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "C";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "D";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "E";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "F";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Unit";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Data Type";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Date/Time";
            for (int i = 0; i < dateTimes.RowCount; i++)
            {
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[i + rowStart, 0].Value = CellToString(dateTimes.Cells[i, 0]);
            }

            
            int colStart = 1;
            for (int i = 0; i < values.ColumnCount; i++)
            {
                rowStart = 0;
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i + colStart].Value = ts_paths[i].Apart;
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i + colStart].Value = ts_paths[i].Bpart;
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i + colStart].Value = ts_paths[i].Cpart;
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i + colStart].Value = "";
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i + colStart].Value = "";
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i + colStart].Value = ts_paths[i].Fpart;
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i + colStart].Value = "Unit";
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i + colStart].Value = "Data Type";
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i + colStart].Value = "Value " + (i + 1).ToString();
                for (int j = 0; j < values.RowCount; j++)
                {
                    ExcelView.ActiveWorkbook.Worksheets[0].Cells[j + rowStart, i + colStart].Value = CellToString(values.Cells[j, i]);
                }
            }
            ExcelView.ActiveWorkbookSet.ReleaseLock();
        }

        private void ShowPairedDataPreview(IRange ordinates, IRange values)
        {
            ExcelView.ActiveWorkbookSet.GetLock();
            ExcelView.ActiveWorksheet.Cells.Clear();

            int rowStart = 0;
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "A";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "B";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "C";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "D";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "E";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "F";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Unit 1";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Unit 2";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Data Type 1";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Data Type 2";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Ordinates";
            for (int i = 0; i < ordinates.RowCount; i++)
            {
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[i + rowStart, 0].Value = CellToString(ordinates.Cells[i, 0]);
            }

            rowStart = 0;
            int colStart = 1;
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, colStart].Value = pd_path.Apart;
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, colStart].Value = pd_path.Bpart;
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, colStart].Value = pd_path.Cpart;
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, colStart].Value = "";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, colStart].Value = "";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, colStart].Value = pd_path.Fpart;
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Unit 1";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Unit 2";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Data Type 1";
            ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, 0].Value = "Data Type 2";
            for (int i = 0; i < values.ColumnCount; i++)
            {
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[rowStart++, i].Value = "Value " + (i + 1).ToString();
                for (int j = 0; j < values.RowCount; j++)
                {
                    ExcelView.ActiveWorkbook.Worksheets[0].Cells[j + rowStart, i + colStart].Value = CellToString(values.Cells[j, i]);
                }
                rowStart--;
            }
            ExcelView.ActiveWorkbookSet.ReleaseLock();

        }

        public void ResetPaths()
        {
            ts_paths.Clear();
            pd_path = new DssPath();
            PreviousPage = null;
        }

        private void ExcelView_ShowError(object sender, SpreadsheetGear.Windows.Controls.ShowErrorEventArgs e)
        {
            e.Handled = true;
        }

    }
}
