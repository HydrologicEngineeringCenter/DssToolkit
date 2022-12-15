using SpreadsheetGear;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Hec.Dss;

namespace DssExcel
{
  /// <summary>
  /// ExcelTimeSeries has methods for reading and writing time series data to excel using the 
  /// format below.  Multiple sereis are supported by adding additional colums D, E, etc.
  /// 
  /// +-------+-----------------+-------------+
  /// |   A   |        B        |      C      |
  /// +-------+-----------------+-------------+
  /// | A     |  watershed      | CARUTHERS C |
  /// | B     |  location       | IVANPAH CA  |
  /// | C     |  parameter      | FLOW        |
  /// | E     |  interval/block |             |
  /// | F     |  version/tag    | USGS        |
  /// | Units |                 | CFS         |
  /// | Type  |                 | INST-VAL    |
  /// | 1     | 31May2020  2300 | 0.0         |
  /// | 2     | 31May2020  2315 | 0.0         |
  /// | 3     | 31May2020  2330 | 0.0         |
  /// | 4     | 31May2020  2345 | 0.0         |
  /// | 5     | 01Jun2020  0000 | 0.0         |
  /// | 6     | 01Jun2020  0015 | 0.0         |
  /// | 7     | 01Jun2020  0030 | 0.0         |
  /// +-------+-----------------+-------------+
  /// 
  /// 
  /// </summary>
  public class ExcelTimeSeries
  {

    private static string[] firstColumn = { "A", "B", "C", "E", "F", "Units", "Type" };
    private static (int r, int c) indexOfWatershed = (0, 2);
    private static (int r, int c) indexOfLocation = (1, 2);
    private static (int r, int c) indexOfParameter = (2, 2);
    private static (int r, int c) indexOfInterval = (3, 2);
    private static (int r, int c) indexOfVersion = (4, 2);
    private static (int r, int c) indexOfUnits = (5, 2);
    private static (int r, int c) indexOfType = (6, 2);
    private static (int r, int c) indexDates = ( 7, 1);
    private static (int r, int c) indexValues = ( 7, 2);
    
    

    /// <summary>
    /// Writes a time series data from primitives  into a worksheet.
    /// The worksheet is formatted in DSSVue compatable format
    /// </summary>
    /// <param name="worksheet">destination for time series data</param>
    /// <param name="dateTimes"></param>
    /// <param name="values"></param>
    /// <param name="SeriesTitles">titles for series (used in DSS C part)</param>
    /// <param name="locationNames">names for the series locations (used in DSS B part)</param>
    public static void Write(IWorksheet worksheet, DateTime[] dateTimes, double[,] values,
                            string[] SeriesTitles, string[] locationNames)
    {
      Hec.Dss.TimeSeries ts = new Hec.Dss.TimeSeries();
      ts.Times=dateTimes;
      if (dateTimes.Length != values.GetLength(0))
        throw new Exception("The list of datetime is a different length than the list of values");

      worksheet.WorkbookSet.GetLock();
      try
      {
        var range = worksheet.Cells;
        range.Clear();
        Excel.WriteArrayDown(range, firstColumn);
        string ePart = "";
        if (ts.IsRegular(dateTimes))
          ePart = "interval:";
        else
          ePart = "block-size:";
        Excel.WriteArrayDown(range[0,1], new string[] { "watershed:", "location:", "parameter:", ePart, "version:" });
        Excel.WriteArrayAcross(range[2, 2], locationNames);
        Excel.WriteArrayAcross(range[3, 2], SeriesTitles);

        string[] intervals= new string[SeriesTitles.Length];
        for (int i = 0; i < intervals.Length; i++)
        {
          intervals[i]= TimeWindow.GetInterval(ts); 
        }
        Excel.WriteArrayAcross(range[3, 2], intervals);
        Excel.WriteSequenceDown(range[7,0],1,dateTimes.Length);

        int rowOffset = indexDates.r;
        for (int i = 0; i < dateTimes.Length; i++)
        {
          
          var dest = range[i+ rowOffset, 1];
          dest.Value = dateTimes[i];
          dest.NumberFormat = "ddMMMyyyy HH:mm:ss";
        }


        int colStart = indexValues.c;
        for (int col = 0; col < values.GetLength(1); col++)
        {
          worksheet.Cells[indexValues.r, col + colStart].Value = SeriesTitles[col];
          for (int rowIndex = 0; rowIndex < values.GetLength(0); rowIndex++)
          {
            worksheet.Cells[rowIndex + rowOffset, col + colStart].Value = values[rowIndex, col];
          }
        }
        worksheet.Cells["A:A"].Columns.AutoFit();
        worksheet.Cells["B:B"].Columns.AutoFit();
      }
      finally
      {
        worksheet.WorkbookSet.ReleaseLock();
      }
    }

    public static TimeSeries[] Read(string excelFileName, string sheetName = "sheet1")
    {
      var workbook = SpreadsheetGear.Factory.GetWorkbook(excelFileName);
      var sheet = workbook.Worksheets[sheetName];
      TimeSeries[] tsList = Read(sheet);

      return tsList;
    }

    private static TimeSeries[] Read(IWorksheet worksheet)
    {
      var cells =worksheet.Range;
      if (!Excel.IsMatchDown(cells, firstColumn))
        return null;

      var usedRange = worksheet.GetUsedRange(true);

      var dateCells = worksheet.Cells[indexDates.r, indexDates.c,usedRange.RowCount-1, indexDates.c];

      if(! Excel.TryGetDateArray(dateCells,out DateTime[] dates,out string errorMessage))
      {
        Logging.WriteError(errorMessage);
        throw new Exception(errorMessage);
      }
      // find how many series by reading first value for each series.
      string[] firstRow = Excel.ReadAcross(worksheet, cells[indexValues.r, indexValues.c]);


      var rval = new List<TimeSeries>(firstRow.Length);
      for (int i = 0; i < firstRow.Length; i++)
      {
        var valueCells = worksheet.Cells[indexValues.r, indexValues.c+i, usedRange.RowCount - 1, indexValues.c+i];
        if(!Excel.TryGetValueArray(valueCells, out double[] values, out errorMessage))
        {
          Logging.WriteError(errorMessage);
        }
        TimeSeries ts = new TimeSeries();
        ts.Times = dates;
        ts.Values = values;
        ts.Units = Excel.CellString(cells[indexOfUnits.r, indexOfUnits.c + i]);
        ts.DataType = Excel.CellString(cells[indexOfType.r, indexOfType.c + i]);
        rval.Add(ts);
      }


      return rval.ToArray();
    }
  }
}

