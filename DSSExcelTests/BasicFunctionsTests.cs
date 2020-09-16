using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using Hec.Dss;

namespace DSSExcelTests
{
    [TestClass]
    public class BasicFunctionsTests
    {
        [TestMethod]
        public void CheckIfExcelSheetIsRegularTimeSeriesWithIndex()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "indexedRegularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.RegularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsRegularTimeSeriesWithNoIndex()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "regularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.RegularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsIrregularTimeSeriesWithIndex()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "indexedIrregularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.IrregularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsIrregularTimeSeriesWithNoIndex()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "irregularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.IrregularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsPairedDataWithIndex()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "indexedPairedData1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.PairedData);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsPairedDataWithNoIndex()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "exportPD1.xls");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.PairedData);

        }

        [TestMethod]
        public void GetRegularTimeSeriesFromExcel1()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "indexedRegularTimeSeries1.xlsx");
            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetRegularTimeSeriesFromExcel2()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "regularTimeSeries1.xlsx");
            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetIrregularTimeSeriesFromExcel1()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "indexedIrregularTimeSeries1.xlsx");
            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetIrregularTimeSeriesFromExcel2()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "irregularTimeSeries1.xlsx");
            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetPairedDataFromExcel1()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "indexedPairedData1.xlsx");
            var pd = de.GetPairedData("Sheet1");
        }

        [TestMethod]
        public void GetPairedDataFromExcel2()
        {
            ExcelReader de = new ExcelReader(TestUtility.BasePath + "pairedData1.xlsx");
            var pd = de.GetPairedData("Sheet1");
        }


    }
}
