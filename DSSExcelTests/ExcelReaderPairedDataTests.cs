using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;
using System.Linq;

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
            Assert.AreEqual(r.CheckType("Sheet1"), Hec.Dss.RecordType.PairedData);
        }

        [TestMethod]
        public void CheckType_PD()
        {
            var filename = TestUtility.BasePath + "pairedData1.xlsx";
            ExcelReader r = new ExcelReader(filename);
            Assert.AreEqual(r.CheckType("Sheet1"), Hec.Dss.RecordType.PairedData);
        }

        [TestMethod]
        public void CheckIndexPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.HasIndex("Sheet1"), true);
        }

        [TestMethod]
        public void GetPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            var ts = r.GetTimeSeries("Sheet1");
            var expected_times = TestUtility.CreateTimeSeriesTimes("5/31/2020  11:00:00 PM", 10, 0, 0, 15);
            Assert.IsTrue(Enumerable.SequenceEqual(ts.Times, expected_times));
        }

        [TestMethod]
        public void CheckPathRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.IsTrue(r.DSSPathExists("Sheet1", 0));
        }

        [TestMethod]
        public void CheckPathPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.IsTrue(r.DSSPathExists("Sheet1", 0));
        }

        [TestMethod]
        public void CheckPathLayoutPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.GetDSSPathLayout("Sheet1"), PathLayout.PathWithoutDPart);
        }

        [TestMethod]
        public void GetPDPath()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            var ts = r.GetTimeSeries("Sheet1");
            Assert.AreEqual(ts.Path.FullPath, @"/CARUTHERS C/IVANPAH CA/FLOW//15Minute/USGS/");
        }

        [TestMethod]
        public void GetDataStartRowPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.DataStartRow("Sheet1"), 9);
        }

        [TestMethod]
        public void GetDataStartRowIndexPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.DataStartRowIndex("Sheet1"), 8);
        }

        [TestMethod]
        public void GetPathEndRowPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.DSSPathEndRow("Sheet1"), 7);
        }

        [TestMethod]
        public void GetPathEndRowIndexPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.DSSPathEndRowIndex("Sheet1"), 6);
        }

        [TestMethod]
        public void GetRowCountPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.RowCount("Sheet1"), 18);
        }

        [TestMethod]
        public void GetColumnCountPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.ColumnCount("Sheet1"), 3);
        }

        [TestMethod]
        public void GetSmallestColumnCountPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.SmallestColumnRowCount("Sheet1"), 18);
        }

        [TestMethod]
        public void IsPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            r.GetTimeSeries("Sheet1");
            Assert.AreEqual(r.isRegularTimeSeries("Sheet1"), true);
        }
    }
}
