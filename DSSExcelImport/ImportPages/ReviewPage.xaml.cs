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


    public void SetupReviewPage(RecordType recordType, IRange range1, IRange range2)
    {
      if (recordType is RecordType.IrregularTimeSeries || recordType is RecordType.RegularTimeSeries)
      {
        currentRecordType = RecordType.RegularTimeSeries;
      }
      else if (recordType is RecordType.PairedData)
      {
        currentRecordType = RecordType.PairedData;
      }
      SetupRecordPreview(recordType, range1, range2);
      ExcelView.ActiveWorkbookSet.GetLock();
      ExcelView.ActiveWorksheet.Cells.Columns.AutoFit();
      ExcelView.ActiveWorkbookSet.ReleaseLock();
    }


    private void SetupRecordPreview(RecordType recordType, IRange range1, IRange range2)
    {
      if (recordType is RecordType.RegularTimeSeries || recordType is RecordType.IrregularTimeSeries)
        ShowTimeSeriesPreview(range1, range2);
      else if (recordType is RecordType.PairedData)
        ShowPairedDataPreview(range1, range2);
    }

    private static string[] EstimateCPart(IRange values)
    {
      List<string> rval = new List<string>();
      for (int c = 0; c < values.ColumnCount; c++)
      {
        var s = "value " + (c + 1);
        if (values.Row > 0)
        {  // look at previous row for column names

          IRange r = values.Cells[-1, c];
          if (r != null)
            s = values.Cells[-1, c].Value.ToString();
        }
         
        rval.Add(s);
      }

      return rval.ToArray();
    }
    private void ShowTimeSeriesPreview(IRange dateTimeRange, IRange values)
    {
      ExcelView.ActiveWorkbookSet.GetLock();
      try
      {
        ExcelView.ActiveWorksheet.Cells.Clear();
        
        var worksheet = ExcelView.ActiveWorkbook.Worksheets[0];
        var range = worksheet.Cells;
        string[] cParts = EstimateCPart(values);

        range[0, 0].Value = "A";
        range[1, 0].Value = "B";
        for (int i = 0; i < cParts.Length; i++)
        {
          range[1, i + 1].Value = values.Worksheet.Workbook.Name;
        }
        range[2, 0].Value = "C";
        for (int i = 0; i < cParts.Length; i++)
        {
          range[2, i + 1].Value = cParts[i];
        }
        range[3, 0].Value = "D";
        range[4, 0].Value = "E";
        range[5, 0].Value = "F";
        range[6, 0].Value = "Unit";
        range[7, 0].Value = "Data Type";

        range[8, 0].Value = "Date/Time";
        int rowIndex = 9;
        for (int i = 0; i < dateTimeRange.RowCount; i++)
        {
          var dest = range[i + rowIndex, 0];
          var src = dateTimeRange.Cells[i, 0];
          dest.Value = src.Value;

          dest.NumberFormat = "ddMMMyyyy HH:mm:ss";
        }


        int colStart = 1;
        for (int i = 0; i < values.ColumnCount; i++)
        {
          worksheet.Cells[8, i + colStart].Value = cParts[i];
          for (int j = 0; j < values.RowCount; j++)
          {
            worksheet.Cells[j + rowIndex, i + colStart].Value =  CellToString(values.Cells[j, i]);
          }
        }
      }
      finally
      {
        ExcelView.ActiveWorkbookSet.ReleaseLock();
      }
    }

    private void ShowPairedDataPreview(IRange ordinates, IRange values)
    {
      ExcelView.ActiveWorkbookSet.GetLock();
      ExcelView.ActiveWorksheet.Cells.Clear();

      int headerEntry = 0;
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "A";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "B";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "C";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "D";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "E";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "F";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "Unit 1";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "Unit 2";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "Data Type 1";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry++, 0].Value = "Data Type 2";
      ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry, 0].Value = "Ordinates";

      int rowStart = headerEntry + 1;
      for (int i = 0; i < ordinates.RowCount; i++)
      {
        ExcelView.ActiveWorkbook.Worksheets[0].Cells[i + rowStart, 0].Value = CellToString(ordinates.Cells[i, 0]);
      }

      int colStart = 1;
      for (int i = 0; i < values.ColumnCount; i++)
      {
        ExcelView.ActiveWorkbook.Worksheets[0].Cells[headerEntry, i + colStart].Value = "Value " + (i + 1).ToString();
        for (int j = 0; j < values.RowCount; j++)
        {
          ExcelView.ActiveWorkbook.Worksheets[0].Cells[j + rowStart, i + colStart].Value = CellToString(values.Cells[j, i]);
        }
      }
      ExcelView.ActiveWorkbookSet.ReleaseLock();

    }

    private void ExcelView_ShowError(object sender, SpreadsheetGear.Windows.Controls.ShowErrorEventArgs e)
    {
      e.Handled = true;
    }

  }
}
