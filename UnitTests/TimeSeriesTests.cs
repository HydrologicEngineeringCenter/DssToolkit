using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using DssExcel;
using Hec.Dss;

namespace UnitTests
{
  [TestClass]
  public class TimeSeriesTests
  {

    /// <summary>
    /// Read time series from Excel.
    /// The excel sheet is in the DssVue export format
    /// </summary>
    [TestMethod]
    public void RegularIntervalFromExcel()
    {
      var filename = TestUtility.BasePath + "indexedRegularTimeSeries1.xlsx";
      var tsList = ExcelTimeSeries.Read(filename);
      Assert.AreEqual(1, tsList.Length);
      CheckTimeSeries(tsList[0]);

      // write to in-memory workbook, and read back for comparision
      var ws = SpreadsheetGear.Factory.GetWorkbook().Worksheets[0];
      ExcelTimeSeries.Write(ws, tsList);

      var tsList2 = ExcelTimeSeries.Read(ws);
      CheckTimeSeries(tsList2[0]);

    }

    private static void CheckTimeSeries(TimeSeries ts)
    {
      Assert.AreEqual(4206, ts.Times.Length);
      Assert.AreEqual(4206, ts.Count);
      Assert.AreEqual("CFS", ts.Units);
      Assert.AreEqual("INST-VAL", ts.DataType);
      Assert.AreEqual("CARUTHERS C", ts.Path.Apart);
      Assert.AreEqual("IVANPAH CA", ts.Path.Bpart);
      Assert.AreEqual("FLOW", ts.Path.Cpart);
      Assert.AreEqual("USGS", ts.Path.Fpart);


    }
  }
}
