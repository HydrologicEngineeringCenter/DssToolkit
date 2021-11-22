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

         Assert.AreEqual(t, new DateTime(2020, 7, 14, 6, 1, 0));
         Assert.AreEqual(t2, new DateTime(2020, 7, 14, 6, 1, 15));


      }
  }
}
