using SpreadsheetGear;

namespace Hec.Excel
{
  public static class ExcelExtensions
  {
    // https://stackoverflow.com/questions/67207520/how-to-get-actual-columns-range-spreadsheetgear
    /// <summary>
    ///
    /// </summary>
    /// <param name="worksheet"></param>
    /// <param name="ignoreEmptyCells"></param>
    /// <returns></returns>
    public static IRange GetUsedRange(this IWorksheet worksheet, bool ignoreEmptyCells)
    {
      IRange usedRange = worksheet.UsedRange;
      if (!ignoreEmptyCells)
        return usedRange;

      // Find last row in used range with a cell containing data.
      IRange foundCell = usedRange.Find("*", usedRange[0, 0], FindLookIn.Formulas,
          LookAt.Part, SearchOrder.ByRows, SearchDirection.Previous, false);
      int lastRow = foundCell?.Row ?? 0;

      // Find last column in used range with a cell containing data.
      foundCell = usedRange.Find("*", usedRange[0, 0], FindLookIn.Formulas,
          LookAt.Part, SearchOrder.ByColumns, SearchDirection.Previous, false);
      int lastCol = foundCell?.Column ?? 0;

      // Return a new used range that clips of any empty rows/cols.
      return worksheet.Cells[worksheet.UsedRange.Row, worksheet.UsedRange.Column, lastRow, lastCol];
    }
  }
}
