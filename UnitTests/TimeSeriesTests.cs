using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using DssExcel;
using Hec.Dss;
using System.Collections.Generic;
using Hec.Excel;

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
      var ts = TestUtility.TimeSeriesFromExcel("indexedRegularTimeSeries1.xlsx");
      CheckTimeSeries(ts);

      // write to in-memory workbook, and read back for comparision
      var ws = SpreadsheetGear.Factory.GetWorkbook().Worksheets[0];
      ExcelTimeSeries.Write(ws, new TimeSeries[] { ts });

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
      Assert.AreEqual(15, ts.TimeSpanInterval().Minutes);
    }
    /// <summary>
    /// </summary>
    [TestMethod]
    public void Simple_IRRegular()
    {
      var ts = TestUtility.TimeSeriesFromExcel("small-ir-ts.xlsx");
      double[] expectedValues = new double[] { -1, 0, 1, 2, 3, 4, 5, 6, 7, 8 };
      var startTime = new DateTime(2020, 5, 31, 23, 0, 0);
      DateTime[] expectedTimes = CreateTimes(startTime, 15, 10);
      // last time is not on consistent interval
      expectedTimes[expectedTimes.Length - 1] = new DateTime(2020, 6, 1, 4, 45, 0);
      Assert.AreEqual(expectedTimes.Length, ts.Count);
      Assert.AreEqual("CARUTHERS C", ts.Path.Apart);
      Assert.AreEqual("IVANPAH CA", ts.Path.Bpart);
      Assert.AreEqual("FLOW", ts.Path.Cpart);
      Assert.AreEqual("USGS", ts.Path.Fpart);
      Assert.AreEqual("CFS", ts.Units);
      Assert.AreEqual("INST-VAL", ts.DataType);
      Assert.IsTrue(!TimeSeries.IsRegular(ts.Times));
      CollectionAssert.AreEqual(expectedTimes, ts.Times);
      CollectionAssert.AreEqual(expectedValues, ts.Values);

     
    }

    private static DateTime[] CreateTimes(DateTime t1,int incrementInMinutes,int count)
    {
      var expectedTimes = new List<DateTime>();
      DateTime t = t1;
      // 15 minute data until the last point (to make irregular)
      for (int i = 0; i < count; i++)
      {
        expectedTimes.Add(t);
        t = t.AddMinutes(incrementInMinutes);
      }
      
      return expectedTimes.ToArray();
    }
    [TestMethod]
    public void Century_IR()
    {
      var ts = TestUtility.TimeSeriesFromExcel("ir-century.xlsx");
      Assert.AreEqual(113, ts.Count);
      Assert.AreEqual("irregular-time-series", ts.Path.Apart);
      Assert.AreEqual("FAIR OAKS CA", ts.Path.Bpart);
      Assert.AreEqual("FLOW-ANNUAL PEAK", ts.Path.Cpart);
      Assert.AreEqual("USGS", ts.Path.Fpart);
      Assert.AreEqual("CFS", ts.Units);
      Assert.AreEqual("INST-VAL", ts.DataType);
      Assert.IsTrue(!TimeSeries.IsRegular(ts.Times));

      //Assert.AreEqual(new DateTime(1862,1,11),ts.Times[0]); // java export gave -1
      Assert.AreEqual(new DateTime(1899,12,30),ts.Times[0]); // java export gave -1
      Assert.AreEqual(318000, ts.Values[0]);

      Assert.AreEqual(new DateTime(2017,2,11), ts.Times[ts.Count - 1]);
      Assert.AreEqual(85400, ts.Values[ts.Count-1]); 

    }
    [TestMethod]
    public void Simple_Regular()
    {
      var ts = TestUtility.TimeSeriesFromExcel("small-r-ts.xlsx");
      double[] expectedValues = new double[] { -1, 0, 1, 2, 3, 4, 5, 6, 7, 8 };
      var startTime = new DateTime(2020, 5, 31, 23, 0, 0);
      DateTime[] expectedTimes = CreateTimes(startTime, 15, 10);
      Assert.AreEqual(expectedTimes.Length, ts.Count);
      Assert.AreEqual("CARUTHERS C", ts.Path.Apart);
      Assert.AreEqual("IVANPAH CA", ts.Path.Bpart);
      Assert.AreEqual("FLOW", ts.Path.Cpart);
      Assert.AreEqual("USGS", ts.Path.Fpart);
      Assert.AreEqual("CFS", ts.Units);
      Assert.AreEqual("INST-VAL", ts.DataType);
      Assert.IsTrue(TimeSeries.IsRegular(ts.Times));
      CollectionAssert.AreEqual(expectedTimes, ts.Times);
      CollectionAssert.AreEqual(expectedValues, ts.Values);
    }

  }
}
