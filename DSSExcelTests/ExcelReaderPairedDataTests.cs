using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;
using System.Linq;
using System.Collections.Generic;

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
        public void CheckIndexPD1()
        {
            var filename = TestUtility.BasePath + "indexedPairedData1.xlsx";
            ExcelReader r = new ExcelReader(filename);
            Assert.AreEqual(r.HasIndex("Sheet1"), true);
        }
        
        [TestMethod]
        public void CheckIndexPD2()
        {
            var filename = TestUtility.BasePath + "pairedData1.xlsx";
            ExcelReader r = new ExcelReader(filename);
            Assert.AreEqual(r.HasIndex("Sheet1"), false);
        }

        [TestMethod]
        public void GetPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            var pd = r.GetPairedData("Sheet1");
            var expected_ordinates = new double[] { 2, 4, 6, 8, 10, 12, 14 };

            var expected_values = new List<double[]>();
            expected_values.Add(new double[] { 10, 11, 12, 13, 14, 15, 16 });
            expected_values.Add(new double[] { 20, 22, 24, 26, 28, 30, 32 });
            expected_values.Add(new double[] { 30, 33, 36, 39, 42, 45, 48 });

            Assert.IsTrue(Enumerable.SequenceEqual(pd.Ordinates, expected_ordinates));
            Assert.IsTrue(Enumerable.SequenceEqual(pd.Values[0], expected_values[0]));
            Assert.IsTrue(Enumerable.SequenceEqual(pd.Values[1], expected_values[1]));
            Assert.IsTrue(Enumerable.SequenceEqual(pd.Values[2], expected_values[2]));
        }

        [TestMethod]
        public void CheckPathLayoutPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.GetDSSPathLayout("Sheet1"), PathLayout.PD_StandardPath);
        }

        [TestMethod]
        public void GetPDPath()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            var pd = r.GetPairedData("Sheet1");
            Assert.AreEqual(pd.Path.FullPath, @"/a/b/c//e/f/");
            Assert.AreEqual(pd.UnitsIndependent, "u1");
            Assert.AreEqual(pd.UnitsDependent, "u2");
            Assert.AreEqual(pd.TypeIndependent, "t1");
            Assert.AreEqual(pd.TypeDependent, "t2");
        }

        [TestMethod]
        public void GetDataStartRowPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.DataStartRow("Sheet1"), 12);
        }

        [TestMethod]
        public void GetDataStartRowIndexPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.DataStartRowIndex("Sheet1"), 11);
        }

        [TestMethod]
        public void GetPathEndRowPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.DSSPathEndRow("Sheet1"), 10);
        }

        [TestMethod]
        public void GetPathEndRowIndexPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.DSSPathEndRowIndex("Sheet1"), 9);
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
            Assert.AreEqual(r.ColumnCount("Sheet1"), 4);
        }

        [TestMethod]
        public void GetSmallestColumnRowCountPD()
        {
            ExcelReader r = new ExcelReader(TestUtility.SimplePDPath);
            Assert.AreEqual(r.SmallestColumnRowCount("Sheet1"), 18);
        }
    }
}
