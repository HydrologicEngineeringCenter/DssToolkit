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
    public void Simple()
    {
      var filename = TestUtility.BasePath + "simple-paired-data.xlsx";
      PairedData pd = ExcelPairedData.Read(filename);

      Assert.IsNotNull(pd);
      Assert.AreEqual("FEET", pd.UnitsIndependent);
      Assert.AreEqual("CFS", pd.UnitsDependent);
      Assert.AreEqual("UNT1", pd.TypeIndependent);
      Assert.AreEqual("UNT2", pd.TypeDependent);

      Assert.AreEqual(34, pd.Values.Count);
      Assert.AreEqual(0, pd.Ordinates[0]);
      Assert.AreEqual(4.80000019073486, pd.Ordinates[1],0.0001);
      Assert.AreEqual(22.7000007629394, pd.Ordinates[pd.Ordinates.Length-1],0.0001);

      Assert.AreEqual(0, pd.Values[0][0]);
      Assert.AreEqual(13600, pd.Values[pd.YCount-1][0]);

    }
    [TestMethod]
    public void MultiColumn()
    {
      var filename = TestUtility.BasePath + "multi-column-paired-data.xlsx";
      PairedData pd = ExcelPairedData.Read(filename);
      string path = "/paired-data-multi-column/RIVERDALE/FREQ-FLOW/MAX ANALYTICAL//1969-01 H33(MAX)/";
      Assert.AreEqual(path, pd.Path.FullPath);
      Assert.AreEqual("COMPUTED", pd.Labels[0]);
      Assert.AreEqual("EXP PROB", pd.Labels[1]);
      Assert.AreEqual("5%LIMIT", pd.Labels[2]);
      Assert.AreEqual("95%LIMIT", pd.Labels[3]);
      Assert.AreEqual("PERCENT", pd.UnitsIndependent);
      Assert.AreEqual("CFS", pd.UnitsDependent);
      Assert.AreEqual("FREQ", pd.TypeIndependent);
      Assert.AreEqual("FLOW", pd.TypeDependent);


    }
  }
}
