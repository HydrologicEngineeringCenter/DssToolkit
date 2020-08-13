using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DSSExcelPlugin;
using System.Linq;
using System.Collections.Generic;
using System.Data;

namespace DSSExcelTests
{
    [TestClass]
    public class BasicFunctionsTests
    {

        public BasicFunctionsTests()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
        }

        [TestMethod]
        public void CheckIfExcelSheetIsRegularTimeSeriesWithIndex()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedRegularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.RegularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsRegularTimeSeriesWithNoIndex()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\regularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.RegularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsIrregularTimeSeriesWithIndex()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedIrregularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.IrregularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsIrregularTimeSeriesWithNoIndex()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\irregularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.IrregularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsPairedDataWithIndex()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedPairedData1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.PairedData);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsPairedDataWithNoIndex()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\pairedData1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.PairedData);

        }

        [TestMethod]
        public void GetRegularTimeSeriesFromExcel1()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedRegularTimeSeries1.xlsx");
            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetRegularTimeSeriesFromExcel2()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\regularTimeSeries1.xlsx");
            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetIrregularTimeSeriesFromExcel1()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedIrregularTimeSeries1.xlsx");
            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetIrregularTimeSeriesFromExcel2()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\irregularTimeSeries1.xlsx");
            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetPairedDataFromExcel1()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedPairedData1.xlsx");
            var pd = de.GetPairedData("Sheet1");
        }

        [TestMethod]
        public void GetPairedDataFromExcel2()
        {
            DSSExcel de = new DSSExcel(@"C:\Temp\pairedData1.xlsx");
            var pd = de.GetPairedData("Sheet1");
        }


    }
}
