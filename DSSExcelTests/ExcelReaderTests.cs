using Hec.Dss.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;

namespace DSSExcelTests
{
  [TestClass]
  public class ExcelReaderTests
  {
    [TestMethod]
    public void ReadDateTimes()
    {
      string excelFileName = TestUtility.BasePath + "DATA for Imports.xls";
      ExcelReader r = new ExcelReader(excelFileName);
      var ws = r.workbook.Worksheets[0];
      Console.WriteLine(ws.Name);

      var rng = ws.Cells[7, 1]; //14Jul2020 06:01:00
      var rng2 = ws.Cells[8, 1];//14Jul2020 06:01:15

         DateTime t =r.GetDateFromCell(rng);
         DateTime t2 = r.GetDateFromCell(rng2);

         Console.WriteLine(rng.Value);
      Console.WriteLine(rng.Text);

      //d.Add(GetDateFromString(CellToString(workbook.Worksheets[worksheet].Cells[i, offset])));

    }
  }
}
