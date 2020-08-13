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

        [TestMethod]
        public void CommandLineImport1()
        {
            var fn = TestUtility.BasePath + "regularTimeSeries1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport1.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            DSSExcel de = new DSSExcel(a[1]);
            de.Import(a[2], 0);
        }

        [TestMethod]
        public void CommandLineImport2()
        {
            var fn = TestUtility.BasePath + "indexedRegularTimeSeries1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport1.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            DSSExcel de = new DSSExcel(a[1]);
            de.Import(a[2], 0);
        }

        [TestMethod]
        public void CommandLineImport3()
        {
            var fn = TestUtility.BasePath + "irregularTimeSeries1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport1.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            DSSExcel de = new DSSExcel(a[1]);
            de.Import(a[2], 0);
        }

        [TestMethod]
        public void CommandLineImport4()
        {
            var fn = TestUtility.BasePath + "indexedIrregularTimeSeries1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport1.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            DSSExcel de = new DSSExcel(a[1]);
            de.Import(a[2], 0);
        }

        [TestMethod]
        public void CommandLineImport5()
        {
            var fn = TestUtility.BasePath + "pairedData1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport1.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            DSSExcel de = new DSSExcel(a[1]);
            de.Import(a[2], 0);
        }

        [TestMethod]
        public void CommandLineImport6()
        {
            var fn = TestUtility.BasePath + "indexedPairedData1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport1.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            DSSExcel de = new DSSExcel(a[1]);
            de.Import(a[2], 0);
        }
    }
}
