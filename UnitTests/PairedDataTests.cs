using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using DssExcel;
using Hec.Dss;

namespace UnitTests
{
  [TestClass]
  public class PairedDataTests
  {
    [TestMethod]
    public void SimplePairedData()
    {
      var filename = TestUtility.BasePath + "simple-paired-data.xlsx";
      PairedData pd = ExcelPairedData.Read(filename);

      Assert.IsNotNull(pd);
      Assert.AreEqual("FEET", pd.UnitsIndependent);
      Assert.AreEqual("CFS", pd.UnitsDependent);
      Assert.AreEqual("UNT1", pd.TypeIndependent);


    }
  }
}
