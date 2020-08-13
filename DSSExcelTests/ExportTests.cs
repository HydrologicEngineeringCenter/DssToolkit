using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using DSSExcelPlugin;
using Hec.Dss;
using System.Data.SqlTypes;
using System.Linq;
using System.IO;

namespace DSSExcelTests
{
    /// <summary>
    /// Summary description for ExportTests
    /// </summary>
    [TestClass]
    public class ExportTests
    {
        public ExportTests()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
        }

        [TestMethod]
        public void ExportRegularTimeSeries()
        {
            var fn = @"C:\Temp\exportRTS1.xls";
            File.Delete(fn);
            TimeSeries ts = TestUtility.CreateTimeSeries(10);
            DSSExcel.Export(fn, ts);
        }

        [TestMethod]
        public void ExportIrregularTimeSeries()
        {
            var fn = @"C:\Temp\exportITS1.xls";
            File.Delete(fn);
            TimeSeries ts = TestUtility.CreateTimeSeries(10, false);
            DSSExcel.Export(fn, ts);

        }

        [TestMethod]
        public void ExportPairedData()
        {
            var fn = @"C:\Temp\exportPD1.xls";
            File.Delete(fn);
            PairedData pd = TestUtility.CreatePairedData(5, 10);
            DSSExcel.Export(fn, pd);
        }

        
    }
}
