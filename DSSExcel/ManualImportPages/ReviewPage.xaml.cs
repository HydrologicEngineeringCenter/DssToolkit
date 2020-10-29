using Hec.Dss;
using SpreadsheetGear;
using Hec.Dss.Excel;
using System.Windows;
using System.Windows.Controls;
using static Hec.Dss.Excel.ExcelTools;

namespace DSSExcel
{
    /// <summary>
    /// Interaction logic for DSSPathPage.xaml
    /// </summary>
    public partial class ReviewPage : UserControl
    {
        public UserControl PreviousPage;
        public RecordType currentRecordType;
        private bool tsPathGenerated = false;
        private bool pdPathGenerated = false;
        private DssPath tsPath = new DssPath();
        private DssPath pdPath = new DssPath();

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


        public string GetPath 
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

        private void ShowTimeSeriesPath()
        {
            if (!tsPathGenerated)
                GenerateTimeSeriesPath();
            DataContext = tsPath;
        }

        private void GenerateTimeSeriesPath()
        {
            tsPath.Apart = "a" + RandomString(3);
            tsPath.Bpart = "b" + RandomString(3);
            tsPath.Cpart = "c" + RandomString(3);
            tsPath.Dpart = "";
            tsPath.Epart = "";
            tsPath.Fpart = "TimeSeries" + RandomString(3);
            tsPathGenerated = true;
        }

        private void ShowPairedDataPath()
        {
            if (!pdPathGenerated)
                GeneratePairedDataPath();
            DataContext = pdPath;
        }

        private void GeneratePairedDataPath()
        {
            pdPath.Apart = "a" + RandomString(3);
            pdPath.Bpart = "b" + RandomString(3);
            pdPath.Cpart = "c" + RandomString(3);
            pdPath.Dpart = "";
            pdPath.Epart = "e" + RandomString(3);
            pdPath.Fpart = "PairedData" + RandomString(3);
            pdPathGenerated = true;
        }

        public void ShowPath(RecordType recordType, IRange range1, IRange range2)
        {
            IsReadOnly(false);
            if (recordType is RecordType.IrregularTimeSeries || recordType is RecordType.RegularTimeSeries)
            {
                currentRecordType = RecordType.RegularTimeSeries;
                ShowTimeSeriesPath();
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
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[0, i].Value = "Values" + i.ToString();
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
                ExcelView.ActiveWorkbook.Worksheets[0].Cells[0, i].Value = "Values" + i.ToString();
                for (int j = 1; j < values.RowCount + 1; j++)
                {
                    ExcelView.ActiveWorkbook.Worksheets[0].Cells[j, i].Value = CellToString(values.Cells[j - 1, i - 1]);
                }
            }
            ExcelView.ActiveWorkbookSet.ReleaseLock();

        }

        public void ResetPath()
        {
            tsPathGenerated = false;
            pdPathGenerated = false;
            tsPath = new DssPath();
            pdPath = new DssPath();
            PreviousPage = null;
        }

        private void ExcelView_ShowError(object sender, SpreadsheetGear.Windows.Controls.ShowErrorEventArgs e)
        {
            e.Handled = true;
        }
    }
}
