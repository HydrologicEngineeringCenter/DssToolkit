using SpreadsheetGear;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hec.Excel
{
  public enum ExcelDirection { Down, Across};
  public class Excel
  {
    string fileName;
    public IWorkbook Workbook { get; set; }
    public Excel(string fileName)
    {
      this.fileName = fileName;
      Workbook = SpreadsheetGear.Factory.GetWorkbook(fileName);
    }
    public static bool TryReadingDouble(IRange r, out double d)
    {
      d = default(double);
      if (r == null)
        return false;
      object o = r.Value;
      
      
      if (r.ValueType == SpreadsheetGear.ValueType.Number)
      {
        d = (double)r.Value;
      }
      else if( r.ValueType == SpreadsheetGear.ValueType.Text)
        return Double.TryParse( r.Text, out d);

      return true;
    }
    public static string RangeToString(IRange cell)
    {
      if (cell == null)
        return "";
      return cell.GetAddress(true, true, ReferenceStyle.A1, true, null);
    }
    /// <summary>
    /// https://github.com/usbr/Pisces/blob/master/TimeSeries.Excel/SpreadsheetGearExcel.cs#L313
    /// </summary>
    /// <param name="d"></param>
    /// <param name="workbook"></param>
    /// <returns></returns>
    private static DateTime DoubleToDateTime(double d, IWorkbook workbook)
    {
      var t = new DateTime();

      //January 1st, 1900 is 1.
      if (d < 0)
      {// hack.. negative number before jan 1,1900
        DateTime j = new DateTime(1900, 1, 1);
        t = j.AddDays(d - 1);
      }
      else
      {
        t = workbook.NumberToDateTime(d); // doesn't support negative values
      }
      return t;
    }
    public static string GetString(IRange cell)
    {
      if( EmptyCell(cell))
        return "";
      return cell.Value.ToString();
    }
    public static bool TryGetDateArray( IRange selection, out DateTime[] dates, out string errorMessage)
    {
      
      if(selection==null || selection.ColumnCount !=1)
      {
        errorMessage = "Please select dates in a single column.";
        dates = null;
        return false;
      }
      
      errorMessage = "";
      dates = new DateTime[selection.RowCount];
      for(int i = 0; i < dates.Length; i++) {
        var cell = selection[i,0];
        if (cell.ValueType == SpreadsheetGear.ValueType.Number)
        {
          var t = DoubleToDateTime((double)cell.Value, selection.Worksheet.Workbook);
          dates[i] = t;
        }
        else
        {
          var txt = cell.Text.Trim();
          if (cell.Value == null || string.IsNullOrEmpty(txt))
          {
            errorMessage = "Found a empty cell, but expected a Date/Time: " + RangeToString(cell);
            return false;
          }
          if (TryParseExcelDateString(txt, out DateTime dt))
          {
            dates[i] = dt;
          }
          else
          {
            errorMessage = "Error reading a date: " + RangeToString(cell);
            return false;
          }
        }

      }
    
     return true;
    }

    internal static void WriteArrayDown(IRange range, string[] data)
    {
      for (int i=0;i<data.Length; i++)
      {
        range[i,0].Value= data[i];
      }
    }
    internal static void WriteArrayDown(IRange range, DateTime[] data)
    {
      for (int i = 0; i < data.Length; i++)
      {
        var cell = range[i,0];
        cell.Value = data[i];
        cell.NumberFormat = "ddMMMyyyy HH:mm:ss";
      }
    }

    internal static void WriteMatrix(IRange range, List<double[]> values)
    {
      if (values.Count == 0 || values[0].Length == 0)
        return;

      for (int col = 0; col < values.Count; col++)
      {
        for (int rowIndex = 0; rowIndex < values[col].Length; rowIndex++)
        {
          range[rowIndex, col].Value = values[col][rowIndex];
        }
      }
    }
    internal static void WriteMatrix(IRange range, double[,] values)
    {
      for (int col = 0; col < values.GetLength(1); col++)
      {
        for (int rowIndex = 0; rowIndex < values.GetLength(0); rowIndex++)
        {
          range[rowIndex , col ].Value = values[rowIndex, col];
        }
      }
    }

    internal static void WriteArrayDown(IRange range, double[] data)
    {
      for (int i = 0; i < data.Length; i++)
      {
        range[i, 0].Value = data[i];
      }
    }
      internal static void WriteArrayAcross(IRange range, string[] data)
    {
      for (int i = 0; i < data.Length; i++)
      {
        range[0,i].Value = data[i];
      }
    }

    internal static void WriteSequenceDown(IRange range, int start, int count, int increment=1)
    {
      int value = start;
      for (int i = 0; i < count; i++)
      {
        range[i, 0].Value = value;
        value += increment;
      }
    }

    /// <summary>
    /// Returns true if the values match the range (moving down a column)
    /// </summary>
    /// <param name="range"></param>
    /// <param name="values"></param>
    /// <returns></returns>
    internal static bool IsMatchDown(IRange range, string[] values)
    {
      for (int i = 0; i < values.Length; i++)
      {
        var obj = range[i, 0].Value;
        if(obj==null || obj.ToString() != values[i]) 
        {
          Logging.WriteError("Expected cell contents '" + values[i] + "'. Found: '" + obj + "'");
          return false; 
        }
      }
      return true;
    }

    public static string RangeTitle(IRange selection, string defaultPrefix = "value")
    {
      return RangeTitles(selection, defaultPrefix)[0];
    }

    /// <summary>
    /// Returns a title for each row in the selection. 
    /// Looks at row above selection for 'names'
    /// </summary>
    /// <param name="selection"></param>
    /// <returns></returns>
    public static string[] RangeTitles(IRange selection, string defaultPrefix="value")
    {
      List<string> rval = new List<string>();
      for (int c = 0; c < selection.ColumnCount; c++)
      {
        var s = defaultPrefix+" " + (c + 1);
        if (selection.Row > 0)
        {  // look at previous row for column names

          IRange r = selection.Cells[-1, c];
          if (r != null)
            s = selection.Cells[-1, c].Value.ToString();
        }

        rval.Add(s);
      }

      return rval.ToArray();
    }

    public static bool TryGetValueArray2D(IRange rangeSelection, out double[,] values, out string errorMessage, double? valueForMissingData = null)
    {
      errorMessage = "";
      values = null;
      if( rangeSelection==null)
      {
        errorMessage = "Please select values ";
        return false;
      }

      values = new double[rangeSelection.RowCount, rangeSelection.ColumnCount];
      for (int columnIndex = 0; columnIndex < rangeSelection.ColumnCount; columnIndex++)
      {
        for (int rowIndex = 0; rowIndex < rangeSelection.RowCount; rowIndex++)
        {
          var cell = rangeSelection[rowIndex, columnIndex];
          bool empty = EmptyCell(cell);

          double d;
          if (empty && valueForMissingData.HasValue)
          {
            d = valueForMissingData.Value;
          }
          else if (empty)
          {
            errorMessage = ErrorMessageEmpty(cell);
            return false;
          }
          else if (!TryReadingDouble(cell, out d))
          {
            errorMessage = ErrorMessageParsingNumber(cell);
            return false;
          }

          values[rowIndex,columnIndex] = d;
        }
      }
      return true;
    }

    /// <summary>
    /// Reads up to rowLmit strings in the first column of the range.
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="range"></param>
    /// <param name="rowLimit"></param>
    /// <returns></returns>
    internal static string[] ReadStringsDown(IWorksheet worksheet, IRange range,int rowLimit=int.MaxValue, bool stopOnEmptyCell=false)
    {
      int maxRows = range.CurrentRegion.RowCount;
      maxRows = Math.Min(maxRows, rowLimit);
      List<String> result = new List<String>();
      for (int i = 0; i < maxRows; i++)
      {
        if (EmptyCell(range[i, 0]))
          break;
        result.Add(range[i,0].Value.ToString());
      }

      return result.ToArray();
    }
    internal static string[] ReadStringsAcross(IWorksheet worksheet, IRange range, bool stopOnEmptyCell)
    {
      int maxColumns = range.CurrentRegion.ColumnCount;
      List<String> result = new List<String>();
      for (int i = 0; i < maxColumns; i++)
      {
        if (EmptyCell(range[0, i]))
        {
          if (stopOnEmptyCell)
            break;
          else
          {
            result.Add("");
          }
        }
        else
        {
          result.Add(range[0, i].Value.ToString());
        }
      }

      return result.ToArray();
    }

    private static bool EmptyCell(IRange cell)
    {
      return cell.Value == null || cell.Text.Trim() == "";
    }
    public static bool TryGetValueArray(IRange rangeSelection, out double[] values, out string errorMessage, double? valueForMissingData=null)
    {
      errorMessage = "";
      values = null;
      if (rangeSelection == null|| rangeSelection.ColumnCount != 1)
      {
        errorMessage = "The selection must be one column. There are " + rangeSelection.ColumnCount + " columns selected";
        return false;
      }

      values = new double[rangeSelection.RowCount];
      for (int i = 0; i < values.Length; i++)
      {
        var cell = rangeSelection[i, 0];
        bool empty = EmptyCell(cell);
        double d;
        if (empty && valueForMissingData.HasValue)
        {
          d = valueForMissingData.Value;
        }
        else if (empty)
        {
          errorMessage = ErrorMessageEmpty(cell);
          return false;
        }
        else if (!TryReadingDouble(cell, out d))
        {
          errorMessage = ErrorMessageParsingNumber(cell);
          return false;
        }

        values[i] = d;
      }
      return true;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="range"></param>
    /// <param name="values">pre allocated array to load from the range</param>
    /// <param name="errorMessage"></param>
    /// <returns></returns>
    internal static bool TryGetValues(IRange range, ExcelDirection direction, double[] values, out string errorMessage)
    {
      errorMessage = "";

      for (int i = 0; i < values.Length; i++)
      {
        IRange cell;
        if (direction == ExcelDirection.Down)
          cell = range[i, 0];
        else if (direction == ExcelDirection.Across)
          cell = range[0, i];
        else
          throw new NotImplementedException(direction.ToString());

        if (EmptyCell(cell))
        {
          errorMessage = ErrorMessageEmpty(cell);
          return false;
        }
        if (!TryReadingDouble(cell,out double d))
        {
          errorMessage = ErrorMessageParsingNumber(cell);
          return false;
        }

        values[i] = d;
      }
      return true;
    }

    private static string ErrorMessageEmpty(IRange cell)
    {
      return "Found a empty cell, but expected a value: " + RangeToString(cell);
    }

    private static string ErrorMessageParsingNumber(IRange cell)
    {
      return "Could not convert '" + cell.Text + "'  to a number: " + RangeToString(cell);
    }

    private static bool TryParseExcelDateString(string s, out DateTime d)
    {
      var rval = false;
      if (s.Contains(" 2400") || s.Contains(" 24:00") || s.Contains(" 24:00:00"))
      {
        string tmp;
        tmp = s.Replace(" 2400", " 0000");
        tmp = tmp.Replace(" 24:00", " 00:00");
        tmp = tmp.Replace(" 24:00:00", " 00:00:00");
        if (!DateTime.TryParse(tmp, out d))
          rval = TryParseAdditionalDateTimeFormats(tmp, out d);
        d = d.AddDays(1);
      }
      else
      {
        rval = DateTime.TryParse(s, out d);
        if (rval)
          return true;
        else
          rval = TryParseAdditionalDateTimeFormats(s, out d);
      }
      return rval;
    }

    private static bool TryParseAdditionalDateTimeFormats(string s, out DateTime d)
    {
      string[] formats =
      {
                "ddMMMyyyy HHmm",
                "ddMMMyyyy HH:mm",
                "ddMMMyyyy HH:mm:ss",
                "ddMMMyyyy  HHmm",
                "ddMMMyyyy  HH:mm",
                "ddMMMyyyy  HH:mm:ss"

            };

      if (DateTime.TryParseExact(s, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out d))
        return true;

      return false;
    }

  }
}
