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
        [TestMethod]
        public void CheckIfExcelSheetIsRegularTimeSeriesWithIndex()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedRegularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.RegularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsRegularTimeSeriesWithNoIndex()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\regularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.RegularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsIrregularTimeSeriesWithIndex()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedIrregularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.IrregularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsIrregularTimeSeriesWithNoIndex()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\irregularTimeSeries1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.IrregularTimeSeries);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsPairedDataWithIndex()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedPairedData1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.PairedData);

        }

        [TestMethod]
        public void CheckIfExcelSheetIsPairedDataWithNoIndex()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\pairedData1.xlsx");
            Assert.AreEqual(de.CheckType("Sheet1"), Hec.Dss.RecordType.PairedData);

        }

        [TestMethod]
        public void GetRegularTimeSeriesFromExcel1()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedRegularTimeSeries1.xlsx");

            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetRegularTimeSeriesFromExcel2()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\regularTimeSeries1.xlsx");

            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetIrregularTimeSeriesFromExcel1()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedIrregularTimeSeries1.xlsx");

            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetIrregularTimeSeriesFromExcel2()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\irregularTimeSeries1.xlsx");

            var ts = de.GetTimeSeries("Sheet1");
        }

        [TestMethod]
        public void GetPairedDataFromExcel1()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedPairedData1.xlsx");

            var pd = de.GetPairedData("Sheet1");
        }

        [TestMethod]
        public void GetPairedDataFromExcel2()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            DSSExcel de = new DSSExcel(@"C:\Temp\pairedData1.xlsx");

            var pd = de.GetPairedData("Sheet1");
        }


    }
}
