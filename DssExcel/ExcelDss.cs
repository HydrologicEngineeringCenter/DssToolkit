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
  /// DssExcel has methods for converting between Excel and DSS
  /// 
  /// </summary>
  internal class ExcelDss
  {

    private static string[] firstColumn = { "A", "B", "C", "E", "F", "Unit", "Type" };
    private static (int r, int c) indexDates = ( 7, 1);

    /// <summary>
    /// Writes a time series data from primitives  into a worksheet.
    /// The worksheet is formatted in DSSVue compatable format
    /// </summary>
    /// <param name="worksheet">destination for time series data</param>
    /// <param name="dateTimes"></param>
    /// <param name="values"></param>
    /// <param name="SeriesTitles">titles for series (used in DSS C part)</param>
    /// <param name="locationNames">names for the series locations (used in DSS B part)</param>
    public static void WriteTimeSeriesToExcel(IWorksheet worksheet, DateTime[] dateTimes, double[,] values,
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


        int colStart = 2;
        for (int col = 0; col < values.GetLength(1); col++)
        {
          worksheet.Cells[8, col + colStart].Value = SeriesTitles[col];
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

    public static void WriteTimeSeriesWorksheetToDSS(IWorksheet worksheet, string dssFileName)
    {
      TimeSeries[] tsList = TimeSeriesFromWorksheet(worksheet);

    }

    private static TimeSeries[] TimeSeriesFromWorksheet(IWorksheet worksheet)
    {
      var ts = new TimeSeries();
      var cells =worksheet.Range;
      if (!Excel.IsMatchDown(cells, firstColumn))
        return null;

      if( Excel.TryGetDateArray(cells[indexDates.r,indexDates.c],out DateTime[] dates,out string errorMessage))
      {
        ts.Times = dates;
      }
      else
      {
        Logging.WriteError(errorMessage);
      }

      var rval = new List<TimeSeries>();

      rval.Add(ts);
      return rval.ToArray();
    }
  }
}

