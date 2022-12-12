using SpreadsheetGear;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  /// <summary>
  /// ExcelTimeSeriesCollection represents a list of time series
  /// in a consistent format for import/export
  /// </summary>
  internal class ExcelTimeSeriesCollectionVM
  {
    IWorksheet worksheet;
    public ExcelTimeSeriesCollectionVM(IWorksheet worksheet)
    {
      this.worksheet = worksheet;
    }

    public void Read(DateTime[] dateTimes, double[,] values,
                            string[] SeriesTitles, string[] locationNames)
    {
      worksheet.WorkbookSet.GetLock();
      try
      {
        var range = worksheet.Cells;
        range.Clear();

        range[0, 0].Value = "A";
        range[1, 0].Value = "B";
        for (int i = 0; i < locationNames.Length; i++)
        {
          range[1, i + 1].Value = locationNames[i];
        }
        range[2, 0].Value = "C";
        for (int i = 0; i < SeriesTitles.Length; i++)
        {
          range[2, i + 1].Value = SeriesTitles[i];
        }
        range[3, 0].Value = "D";
        range[4, 0].Value = "E";
        range[5, 0].Value = "F";
        range[6, 0].Value = "Unit";
        range[7, 0].Value = "Data Type";

        range[8, 0].Value = "Date/Time";
        int rowOffset = 9;
        for (int i = 0; i < dateTimes.Length; i++)
        {
          var dest = range[i + rowOffset, 0];
          dest.Value = dateTimes[i];
          dest.NumberFormat = "ddMMMyyyy HH:mm:ss";
        }


        int colStart = 1;
        for (int col = 0; col < values.Rank; col++)
        {
          worksheet.Cells[8, col + colStart].Value = SeriesTitles[col];
          for (int rowIndex = 0; rowIndex < values.GetLength(0); rowIndex++)
          {
            worksheet.Cells[rowIndex + rowOffset, col + colStart].Value = values[rowIndex, col];
          }
        }
        worksheet.Cells["A:A"].Columns.AutoFit();
      }
      finally
      {
        worksheet.WorkbookSet.ReleaseLock();
      }
    }
  }
}

