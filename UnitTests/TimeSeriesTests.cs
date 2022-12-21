using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using DssExcel;

namespace UnitTests
{
  [TestClass]
  public class TimeSeriesTests
  {
    [TestMethod]
    public void RegularFromExcel()
    {
      var filename = TestUtility.BasePath + "indexedRegularTimeSeries1.xlsx";
      var tsList = ExcelTimeSeries.Read(filename);
     Assert.AreEqual(1,tsList.Length);
      var ts = tsList[0];
      Assert.AreEqual(4206, ts.Times.Length);
      Assert.AreEqual(4206, ts.Count);
      Assert.AreEqual("CFS", ts.Units);
      Assert.AreEqual("INST-VAL", ts.DataType);
    }


  }
}
