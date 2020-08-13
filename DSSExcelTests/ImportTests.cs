using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DSSExcelPlugin;
using System.IO;

namespace DSSExcelTests
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class ImportTests
    {
        public ImportTests()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
        }

        [TestMethod]
        public void ImportRegularTimeSeries1()
        {
            File.Delete(@"C:\Temp\indexedRegularTimeSeries1.dss");
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedRegularTimeSeries1.xlsx");
            de.Import(@"C:\Temp\indexedRegularTimeSeries1.dss", "Sheet1");

        }

        [TestMethod]
        public void ImportIrregularTimeSeries1()
        {
            File.Delete(@"C:\Temp\indexedIrregularTimeSeries1.dss");
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedIrregularTimeSeries1.xlsx");
            de.Import(@"C:\Temp\indexedIrregularTimeSeries1.dss", "Sheet1");

        }

        [TestMethod]
        public void ImportPairedData1()
        {
            File.Delete(@"C:\Temp\indexedPairedData1.dss");
            DSSExcel de = new DSSExcel(@"C:\Temp\indexedPairedData1.xlsx");
            de.Import(@"C:\Temp\indexedPairedData1.dss", "Sheet1");

        }

        [TestMethod]
        public void ImportRegularTimeSeries2()
        {
            File.Delete(@"C:\Temp\regularTimeSeries1.dss");
            DSSExcel de = new DSSExcel(@"C:\Temp\regularTimeSeries1.xlsx");
            de.Import(@"C:\Temp\regularTimeSeries1.dss", "Sheet1");

        }

        [TestMethod]
        public void ImportIrregularTimeSeries2()
        {
            File.Delete(@"C:\Temp\irregularTimeSeries1.dss");
            DSSExcel de = new DSSExcel(@"C:\Temp\irregularTimeSeries1.xlsx");
            de.Import(@"C:\Temp\irregularTimeSeries1.dss", "Sheet1");

        }

        [TestMethod]
        public void ImportPairedData2()
        {
            File.Delete(@"C:\Temp\pairedData1.dss");
            DSSExcel de = new DSSExcel(@"C:\Temp\pairedData1.xlsx");
            de.Import(@"C:\Temp\pairedData1.dss", "Sheet1");
        }
    }
}
