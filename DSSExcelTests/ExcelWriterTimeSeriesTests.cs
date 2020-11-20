using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;
using System.IO;
using Hec.Dss;
using System.Linq;

namespace DSSExcelTests
{
    [TestClass]
    public class ExcelWriterTimeSeriesTests
    {
        [TestMethod]
        public void CreateWorkbookWithExistingFile()
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
        public void CreateWorkbookWithNonExistingFile()
        {
            var filename = TestUtility.OutputPath + "writer-test-non-existing.xlsx";
            File.Delete(filename);
            TimeSeries ts1 = TestUtility.CreateTimeSeries(10);
            ExcelWriter w = new ExcelWriter(filename);
            w.Write(ts1, "Sheet1");

            ExcelReader r = new ExcelReader(filename);
            var ts2 = r.GetTimeSeries("Sheet1");

            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Times, ts2.Times));
            Assert.IsTrue(Enumerable.SequenceEqual(ts1.Values, ts2.Values));
        }

        [TestMethod]
        public void WriteTest()
        {

        }

        [TestMethod]
        public void WriteMultipleTest()
        {

        }

        [TestMethod]
        public void ClearSheetTest()
        {

        }

        [TestMethod]
        public void AddSheetTest()
        {

        }

        [TestMethod]
        public void SheetExistTest1()
        {

        }

        [TestMethod]
        public void SheetExistTest2()
        {

        }
    }
}
