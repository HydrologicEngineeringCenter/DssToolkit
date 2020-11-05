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

        [TestMethod]
        public void CheckIndexRegularTS()
        {

        }

        [TestMethod]
        public void CheckIndexIrregularTS()
        {

        }

        [TestMethod]
        public void GetRegularTSTimes()
        {

        }

        [TestMethod]
        public void GetIrregularTSTimes()
        {

        }

        [TestMethod]
        public void GetRegularTSValues()
        {

        }

        [TestMethod]
        public void GetIrregularTSValues()
        {

        }

        [TestMethod]
        public void CheckPathRegularTS()
        {

        }

        [TestMethod]
        public void CheckPathIrregularTS()
        {

        }

        [TestMethod]
        public void CheckPathLayoutRegularTS()
        {

        }

        [TestMethod]
        public void CheckPathLayoutIrregularTS()
        {

        }

        [TestMethod]
        public void GetRegularTSPath()
        {

        }

        [TestMethod]
        public void GetIrregularTSPath()
        {

        }

        [TestMethod]
        public void GetRegularTS()
        {

        }

        [TestMethod]
        public void GetIrregularTS()
        {

        }

        [TestMethod]
        public void GetDataStartRowRegularTS()
        {

        }

        [TestMethod]
        public void GetDataStartRowIrregularTS()
        {

        }

        [TestMethod]
        public void GetDataStartRowIndexRegularTS()
        {

        }

        [TestMethod]
        public void GetDataStartRowIndexIrregularTS()
        {

        }

        [TestMethod]
        public void GetPathEndRowRegularTS()
        {

        }

        [TestMethod]
        public void GetPathEndRowIrregularTS()
        {

        }

        [TestMethod]
        public void GetPathEndRowIndexRegularTS()
        {

        }

        [TestMethod]
        public void GetPathEndRowIndexIrregularTS()
        {

        }

        [TestMethod]
        public void GetRowCountRegularTS()
        {

        }

        [TestMethod]
        public void GetRowCountIrregularTS()
        {

        }

        [TestMethod]
        public void GetColumnCountRegularTS()
        {

        }

        [TestMethod]
        public void GetColumnCountIrregularTS()
        {

        }

        [TestMethod]
        public void GetSmallestColumnCountRegularTS()
        {

        }

        [TestMethod]
        public void GetSmallestColumnCountIrregularTS()
        {

        }

        [TestMethod]
        public void IsRegularTS()
        {

        }

        [TestMethod]
        public void IsIrregularTS()
        {

        }

    }
}
