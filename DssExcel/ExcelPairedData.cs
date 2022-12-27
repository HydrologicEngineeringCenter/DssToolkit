using System;
using Hec.Dss;
using SpreadsheetGear;

namespace DssExcel
{
  /**
   * ExcelPairedData has methods for reading and writing paired-data to/from excel using the 
  format below.    (The first row (path) is optional)
+--------+---------------------------------------------------------------------------------+
|  Path  | /paired-data-multi-column/RIVERDALE/FREQ-FLOW/MAX ANALYTICAL//1969-01 H33(MAX)/ |
+--------+---------------------------------------------------------------------------------+
| Labels |                      | COMPUTED | EXP PROB |  5%LIMIT | 95%LIMIT |
| Units  | PERCENT              | CFS      |          |          |          |
| Type   | FREQ                 | FLOW     |          |          |          |
| 1      | 0                    | 20912.98 | 24993.01 | 30978.08 | 16020.89 |
| 2      | 0                    | 18679.56 | 21581.48 | 26919.98 | 14562.59 |
| 3      | 1                    | 15956.85 | 17745.02 | 22142.58 | 12741.07 |
| 4      | 1                    | 14054.34 | 15239.98 | 18927.51 | 11434.04 |
| 5      | 2                    | 12272.15 | 13022.79 | 16019.20 | 10178.26 |
| 6      | 5                    | 10074.43 | 10438.13 | 12590.91 | 8576.13  |
| 7      | 10                   | 8505.39  | 8692.51  | 10270.00 | 7383.06  |
| 8      | 20                   | 6981.07  | 7057.34  | 8142.02  | 6165.35  |
| 9      | 50                   | 4893.24  | 4893.24  | 5507.09  | 4340.02  |
| 10     | 80                   | 3529.99  | 3499.13  | 4000.26  | 3021.05  |
| 11     | 90                   | 3009.38  | 2960.99  | 3457.57  | 2504.77  |
| 12     | 95                   | 2652.78  | 2586.76  | 3090.23  | 2153.57  |
+--------+------------------------------------------------------------------+
 
   */
  public class ExcelPairedData
  {
    private static string[] firstColumn = { "Path", "Labels", "Units", "Type" };
    private static (int r, int c) indexOfPath = (0, 1);
    private static (int r, int c) indexOfLabels = (1, 2);
    private static (int r, int c) indexOfUnits = (2, 1);
    private static (int r, int c) indexOfType = (3, 1);
    private static (int r, int c) indexOfData = (4, 0);
    private static (int r, int c) indexOfX = (4, 1);
    private static (int r, int c) indexOfY = (4, 2);

    public static PairedData Read(string excelFileName, string sheetName = "Sheet1")
    {
      var workbook = SpreadsheetGear.Factory.GetWorkbook(excelFileName);
      var sheet = workbook.Worksheets[sheetName];
      PairedData pd = Read(sheet);

      return pd;

    }

    private static PairedData Read(IWorksheet worksheet)
    {
      PairedData rval = new PairedData();
      worksheet.WorkbookSet.GetLock();
      try
      {
        var range = worksheet.Cells;
          if (!Excel.IsMatchDown(range, firstColumn))
            return rval;

          rval.Path = new DssPath(ReadPath(range));



      }
      finally
      {
        worksheet.WorkbookSet.ReleaseLock();
      }
      return rval;
    }

    private static string ReadPath(IRange range)
    {
      if(Excel.CellString(range).ToLower() == "path")
      {
        return Excel.CellString(range[indexOfPath.r, indexOfPath.c]);
      }
      return "";
    }

    public static void Write(IWorksheet worksheet, string path, double[] Xvalues, double[,] Yvalues,
                    string xUnits, string yUnits, string xType, string yType, string[] curveLabels)
    {
      worksheet.WorkbookSet.GetLock();
      try
      {
        var range = worksheet.Cells;
        range.Clear();
        Excel.WriteArrayDown(range[0, 0], firstColumn);
        range[indexOfPath.r, indexOfPath.c].Value = path;
        Excel.WriteArrayAcross(range[indexOfLabels.r, indexOfLabels.c], curveLabels);
        Excel.WriteArrayAcross(range[indexOfUnits.r, indexOfUnits.c], new string[] { xType, yUnits });
        Excel.WriteArrayAcross(range[indexOfType.r, indexOfType.c], new string[] { xType, yType });
        Excel.WriteSequenceDown(range[indexOfData.r, indexOfData.c], 1, Xvalues.Length);
        Excel.WriteArrayDown(range[indexOfX.r, indexOfX.c], Xvalues);
        Excel.WriteMatrix(range[indexOfY.r, indexOfY.c], Yvalues);

      }
      finally
      {
        worksheet.WorkbookSet.ReleaseLock();
      }
    } 
  }
  
}
