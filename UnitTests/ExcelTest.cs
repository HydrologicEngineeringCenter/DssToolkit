using Hec.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;

namespace UnitTests
{
  [TestClass]
  public class ExcelTest
  {

    [TestMethod]
    public void ReadExcelWithBlanks()
    {

      Excel e = new Excel(Path.Combine(TestUtility.BasePath, "time-series-with-blanks.csv"));
      var range = e.Workbook.ActiveSheet.EvaluateRange("G2:H14");
      double missingValue = -123.123;
      if (!Excel.TryGetValueArray2D(range, out double[,] values, out string error, missingValue))
      {
        throw new Exception(error);
      }

      Assert.AreEqual(14.09, values[0, 0]);
      Assert.AreEqual(missingValue, values[4, 0]);
      Assert.AreEqual(missingValue, values[7, 1]);
      Assert.AreEqual(8.72, values[12, 0]);
      Assert.AreEqual(2342, values[12, 1]);

    }

  }
}
