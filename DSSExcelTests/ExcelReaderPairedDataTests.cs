using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;

namespace DSSExcelTests
{
    [TestClass]
    public class ExcelReaderPairedDataTests
    {
        [TestMethod]
        public void CheckType_IndexedPD()
        {
            var filename = TestUtility.BasePath + "indexedPairedData1.xlsx";
            ExcelReader r = new ExcelReader(filename);
            Assert.AreEqual(r.CheckType(0), Hec.Dss.RecordType.PairedData);
        }

        [TestMethod]
        public void CheckType_PD()
        {
            var filename = TestUtility.BasePath + "pairedData1.xlsx";
            ExcelReader r = new ExcelReader(filename);
            Assert.AreEqual(r.CheckType(0), Hec.Dss.RecordType.PairedData);
        }
    }
}
