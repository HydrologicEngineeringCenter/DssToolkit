using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;
using System.IO;
using Hec.Dss;
using System.Linq;
using System.Collections.Generic;

namespace DSSExcelTests
{
    [TestClass]
    public class ExcelWriterTimeSeriesTests
    {
        [TestMethod]
        public void InstantiateExcelWriterWithExistingFile()
        {
            var filename = TestUtility.BasePath + "write-test-existing.xlsx";
            ExcelWriter w = new ExcelWriter(filename);
        }
        
        [TestMethod]
        public void InstantiateExcelWriterWithNonExistingFile()
        {
            var filename = TestUtility.OutputPath + "writer-test-non-existing.xlsx";
            File.Delete(filename);
            ExcelWriter w = new ExcelWriter(filename);
        }

        [TestMethod]
        public void WriteTest()
        {
            var filename = TestUtility.BasePath + "write-test-existing.xlsx";
            TimeSeries ts1 = TestUtility.CreateTimeSeries(10);
            ExcelWriter w = new ExcelWriter(filename);
            w.Write(ts1, "Sheet1");

            ExcelReader r = new ExcelReader(filename);
            var ts2 = r.GetTimeSeries("Sheet1");

            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Times, ts2.Times));
            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Values, ts2.Values));

            w.ClearSheet("Sheet1");
        }

        [TestMethod]
        public void WriteMultipleTest()
        {
            var filename = TestUtility.BasePath + "write-test-existing.xlsx";
            TimeSeries ts1 = TestUtility.CreateTimeSeries(15);
            TimeSeries ts2 = TestUtility.CreateTimeSeries(15);
            ExcelWriter w = new ExcelWriter(filename);
            List<TimeSeries> list1 = new List<TimeSeries>();
            list1.Add(ts1); list1.Add(ts2);
            w.Write(list1, "Sheet1");

            ExcelReader r = new ExcelReader(filename);
            List<TimeSeries> list2 = (List<TimeSeries>)r.GetMultipleTimeSeries("Sheet1");
            TimeSeries ts3 = list2[0];
            TimeSeries ts4 = list2[1];

            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Times, ts3.Times));
            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Values, ts3.Values));
            Assert.IsTrue(Enumerable.SequenceEqual(ts2.Times, ts4.Times));
            Assert.IsTrue(Enumerable.SequenceEqual(ts2.Values, ts4.Values));
        }

        [Ignore]
        [TestMethod]
        public void ClearSheetTest()
        {
            var filename = TestUtility.BasePath + "write-test-existing.xlsx";
            TimeSeries ts1 = TestUtility.CreateTimeSeries(10);
            ExcelWriter w = new ExcelWriter(filename);
            w.Write(ts1, "Sheet1");

            ExcelReader r = new ExcelReader(filename);
            TimeSeries ts2 = r.GetTimeSeries("Sheet1");

            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Times, ts2.Times));
            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Values, ts2.Values));

            w.ClearSheet("Sheet1");

            TimeSeries ts3 = r.GetTimeSeries("Sheet1");

            Assert.AreEqual(new TimeSeries(), ts3);
        }

        [TestMethod]
        public void AddSheetTest()
        {
            var filename = TestUtility.OutputPath + "add-sheet.xlsx";
            var sheetname = "Sheet2";
            File.Delete(filename);

            TimeSeries ts1 = TestUtility.CreateTimeSeries(10);
            ExcelWriter w = new ExcelWriter(filename);
            w.AddSheet(sheetname);
            w.Write(ts1, sheetname);

            ExcelReader r = new ExcelReader(filename);
            var ts2 = r.GetTimeSeries(sheetname);

            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Times, ts2.Times));
            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Values, ts2.Values));

            File.Delete(filename);
        }

        [TestMethod]
        public void SheetExistTest1()
        {
            var filename = TestUtility.BasePath + "write-test-existing.xlsx";
            TimeSeries ts1 = TestUtility.CreateTimeSeries(10);
            ExcelWriter w = new ExcelWriter(filename);
            Assert.IsTrue(w.SheetExists("Sheet1"));
        }

        [TestMethod]
        public void SheetExistTest2()
        {
            var filename = TestUtility.BasePath + "write-test-existing.xlsx";
            TimeSeries ts1 = TestUtility.CreateTimeSeries(10);
            ExcelWriter w = new ExcelWriter(filename);
            Assert.IsFalse(w.SheetExists("Sheet10"));
        }
    }
}
