using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;
using Hec.Dss;
using System.Collections.Generic;
using System.Linq;

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
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.AreEqual(r.HasIndex("Sheet1"), true);
        }

        [TestMethod]
        public void CheckIndexIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.AreEqual(r.HasIndex("Sheet1"), true);
        }

        [TestMethod]
        public void GetRegularTSTimes()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            var ts = r.GetTimeSeries("Sheet1");
            var expected_times = TestUtility.CreateTimeSeriesTimes("5/31/2020  11:00:00 PM", 10, 0, 0, 15);
            Assert.IsTrue(Enumerable.SequenceEqual(ts.Times, expected_times));
        }

        [TestMethod]
        public void GetIrregularTSTimes()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            var ts = r.GetTimeSeries("Sheet1");
            var expected_times = TestUtility.CreateTimeSeriesTimes("5/31/2020  11:00:00 PM", 9, 0, 0, 15);
            expected_times.Add(DateTime.Parse("6/1/2020  4:45:00 AM"));
            Assert.IsTrue(Enumerable.SequenceEqual(ts.Times, expected_times.ToArray()));
        }

        [TestMethod]
        public void GetRegularTSValues()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            var ts = r.GetTimeSeries("Sheet1");
            var expected_values = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            Assert.IsTrue(Enumerable.SequenceEqual(ts.Values, expected_values));
        }

        [TestMethod]
        public void GetIrregularTSValues()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            var ts = r.GetTimeSeries("Sheet1");
            double[] expected_values = new double[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 };
            Assert.IsTrue(Enumerable.SequenceEqual(ts.Values, expected_values));
        }

        [TestMethod]
        public void CheckPathRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.IsTrue(r.DSSPathExists("Sheet1", 0));
        }

        [TestMethod]
        public void CheckPathIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.IsTrue(r.DSSPathExists("Sheet1", 0));
        }

        [TestMethod]
        public void CheckPathLayoutRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.AreEqual(r.GetDSSPathLayout("Sheet1"), PathLayout.TS_PathWithoutDPart);
        }

        [TestMethod]
        public void CheckPathLayoutIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.AreEqual(r.GetDSSPathLayout("Sheet1"), PathLayout.TS_PathWithoutDPart);
        }

        [TestMethod]
        public void GetRegularTSPath()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            var ts = r.GetTimeSeries("Sheet1");
            Assert.AreEqual(ts.Path.FullPath, @"/CARUTHERS C/IVANPAH CA/FLOW//15Minute/USGS/");
        }

        [TestMethod]
        public void GetIrregularTSPath()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            var ts = r.GetTimeSeries("Sheet1");
            Assert.AreEqual(ts.Path.FullPath, @"/CARUTHERS C/IVANPAH CA/FLOW//IR-Year/USGS/");
        }

        [TestMethod]
        public void GetDataStartRowRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.AreEqual(r.DataStartRow("Sheet1"), 9);
        }

        [TestMethod]
        public void GetDataStartRowIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.AreEqual(r.DataStartRow("Sheet1"), 9);
        }

        [TestMethod]
        public void GetDataStartRowIndexRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.AreEqual(r.DataStartRowIndex("Sheet1"), 8);
        }

        [TestMethod]
        public void GetDataStartRowIndexIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.AreEqual(r.DataStartRowIndex("Sheet1"), 8);
        }

        [TestMethod]
        public void GetPathEndRowRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.AreEqual(r.DSSPathEndRow("Sheet1"), 7);
        }

        [TestMethod]
        public void GetPathEndRowIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.AreEqual(r.DSSPathEndRow("Sheet1"), 7);
        }

        [TestMethod]
        public void GetPathEndRowIndexRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.AreEqual(r.DSSPathEndRowIndex("Sheet1"), 6);
        }

        [TestMethod]
        public void GetPathEndRowIndexIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.AreEqual(r.DSSPathEndRowIndex("Sheet1"), 6);
        }

        [TestMethod]
        public void GetRowCountRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.AreEqual(r.RowCount("Sheet1"), 18);
        }

        [TestMethod]
        public void GetRowCountIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.AreEqual(r.RowCount("Sheet1"), 18);
        }

        [TestMethod]
        public void GetColumnCountRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.AreEqual(r.ColumnCount("Sheet1"), 3);
        }

        [TestMethod]
        public void GetColumnCountIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.AreEqual(r.ColumnCount("Sheet1"), 3);
        }

        [TestMethod]
        public void GetSmallestColumnCountRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            Assert.AreEqual(r.SmallestColumnRowCount("Sheet1"), 18);
        }

        [TestMethod]
        public void GetSmallestColumnCountIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            Assert.AreEqual(r.SmallestColumnRowCount("Sheet1"), 18);
        }

        [TestMethod]
        public void IsRegularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleRegularTSPath);
            r.GetTimeSeries("Sheet1");
            Assert.AreEqual(r.isRegularTimeSeries("Sheet1"), true);
        }

        [TestMethod]
        public void IsIrregularTS()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimpleIrregularTSPath);
            r.GetTimeSeries("Sheet1");
            Assert.AreEqual(r.isRegularTimeSeries("Sheet1"), false);
        }

    }
}
