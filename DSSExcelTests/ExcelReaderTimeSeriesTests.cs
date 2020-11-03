using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;

namespace DSSExcelTests
{
    [TestClass]
    public class ExcelReaderTimeSeriesTests
    {
        [TestMethod]
        public void CheckType_IndexedRegularTS()
        {
            var filename = TestUtility.BasePath + "indexedRegularTimeSeries1.xlsx";
            ExcelReader r = new ExcelReader(filename);
            Assert.AreEqual(r.CheckType(0), Hec.Dss.RecordType.RegularTimeSeries);
        }

        [TestMethod]
        public void CheckType_RegularTS()
        {
            var filename = TestUtility.BasePath + "regularTimeSeries1.xlsx";
            ExcelReader r = new ExcelReader(filename);
            Assert.AreEqual(r.CheckType(0), Hec.Dss.RecordType.RegularTimeSeries);
        }

        [TestMethod]
        public void CheckType_IndexedIrregularTS()
        {
            var filename = TestUtility.BasePath + "indexedIrregularTimeSeries1.xlsx";
            ExcelReader r = new ExcelReader(filename);
            Assert.AreEqual(r.CheckType(0), Hec.Dss.RecordType.IrregularTimeSeries);
        }

        [TestMethod]
        public void CheckType_IrregularTS()
        {
            var filename = TestUtility.BasePath + "irregularTimeSeries1.xlsx";
            ExcelReader r = new ExcelReader(filename);
            Assert.AreEqual(r.CheckType(0), Hec.Dss.RecordType.IrregularTimeSeries);
        }
    }
}
