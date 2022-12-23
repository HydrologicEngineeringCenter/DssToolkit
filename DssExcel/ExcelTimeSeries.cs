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
  /// format below.  Multiple series are supported by adding additional colums D, E, etc.
  /// The row labeled D is intentionally skipped
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
                            string[] SeriesTitles, string[] locationNames, string[] versionTags)
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
        Excel.WriteArrayDown(range[0,1], new string[] { "watershed:", "location:", "parameter:", ePart, "version:","units (cfs,feet,...):","  type(PER-AVER,PER-CUM,INST-VAL,INST-CUM):" });
        Excel.WriteArrayAcross(range[indexOfLocation.r, indexOfLocation.c],locationNames);
        Excel.WriteArrayAcross(range[indexOfParameter.r, indexOfParameter.c], SeriesTitles);
        Excel.WriteArrayAcross(range[indexOfVersion.r, indexOfVersion.c], versionTags);

        string[] intervals= new string[SeriesTitles.Length];
        for (int i = 0; i < intervals.Length; i++)
        {
          intervals[i]= TimeWindow.GetInterval(ts); 
        }
        Excel.WriteArrayAcross(range[indexOfInterval.r, indexOfInterval.c], intervals);
        Excel.WriteSequenceDown(range[indexDates.r, 0],1,dateTimes.Length);

        Excel.WriteArrayDown(range[indexDates.r,indexDates.c],dateTimes);

        Excel.WriteMatrix(range[indexValues.r,indexValues.c],values);
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

    public static TimeSeries[] Read(IWorksheet worksheet)
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
      string[] firstRow = Excel.ReadStringsAcross(worksheet, cells[indexValues.r, indexValues.c],true);


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
        ts.Path = GetDssPath(cells[indexOfWatershed.r, indexOfWatershed.c + i]);
        rval.Add(ts);
      }


      return rval.ToArray();
    }

    private static DssPath GetDssPath(IRange range)
    {
      var a = Excel.CellString(range[0, 0]);
      var b = Excel.CellString(range[1, 0]);
      var c = Excel.CellString(range[2, 0]);
      var d = "";
      var e = Excel.CellString(range[3, 0]);
      var f = Excel.CellString(range[4, 0]);
      DssPath p = new DssPath(a, b, c, d, e, f);

      return p;
    }
  }
}

