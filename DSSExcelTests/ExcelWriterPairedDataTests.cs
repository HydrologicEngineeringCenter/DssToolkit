using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;
using Hec.Dss;
using System.Linq;
using System.IO;

namespace DSSExcelTests
{
    [TestClass]
    public class ExcelWriterPairedDataTests
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
            PairedData pd1 = TestUtility.CreatePairedData(3, 10);
            ExcelWriter w = new ExcelWriter(filename);
            w.Write(pd1, "Sheet1");

            ExcelReader r = new ExcelReader(filename);
            PairedData pd2 = r.GetPairedData("Sheet1");

            Assert.IsTrue(Enumerable.SequenceEqual(pd1.Ordinates, pd2.Ordinates));
            Assert.IsTrue(Enumerable.SequenceEqual(pd1.Values[0], pd2.Values[0]));
            Assert.IsTrue(Enumerable.SequenceEqual(pd1.Values[1], pd2.Values[1]));
            Assert.IsTrue(Enumerable.SequenceEqual(pd1.Values[2], pd2.Values[2]));

            w.ClearSheet("Sheet1");
        }

        [Ignore]
        [TestMethod]
        public void ClearSheetTest()
        {
            var filename = TestUtility.BasePath + "write-test-existing.xlsx";
            PairedData pd1 = TestUtility.CreatePairedData(3, 10);
            ExcelWriter w = new ExcelWriter(filename);
            w.Write(pd1, "Sheet1");

            ExcelReader r = new ExcelReader(filename);
            PairedData pd2 = r.GetPairedData("Sheet1");

            Assert.IsTrue(Enumerable.SequenceEqual(pd1.Ordinates, pd2.Ordinates));
            Assert.IsTrue(Enumerable.SequenceEqual(pd1.Values[0], pd2.Values[0]));
            Assert.IsTrue(Enumerable.SequenceEqual(pd1.Values[1], pd2.Values[1]));
            Assert.IsTrue(Enumerable.SequenceEqual(pd1.Values[2], pd2.Values[2]));

            w.ClearSheet("Sheet1");

            PairedData pd3 = r.GetPairedData("Sheet1");

            Assert.AreEqual(new PairedData(), pd3);
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
            ExcelWriter w = new ExcelWriter(filename);
            Assert.IsTrue(w.SheetExists("Sheet1"));
        }

        [TestMethod]
        public void SheetExistTest2()
        {
            var filename = TestUtility.BasePath + "write-test-existing.xlsx";
            ExcelWriter w = new ExcelWriter(filename);
            Assert.IsFalse(w.SheetExists("Sheet10"));
        }
    }
}
