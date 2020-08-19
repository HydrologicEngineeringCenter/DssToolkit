using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;
using System.IO;
using Hec.Dss;

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
            ExcelReader de = new ExcelReader(@"C:\Temp\indexedRegularTimeSeries1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportIrregularTimeSeries1()
        {
            File.Delete(@"C:\Temp\indexedIrregularTimeSeries1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\indexedIrregularTimeSeries1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportPairedData1()
        {
            File.Delete(@"C:\Temp\indexedPairedData1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\indexedPairedData1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportRegularTimeSeries2()
        {
            File.Delete(@"C:\Temp\regularTimeSeries1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\regularTimeSeries1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportIrregularTimeSeries2()
        {
            File.Delete(@"C:\Temp\irregularTimeSeries1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\irregularTimeSeries1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportPairedData2()
        {
            File.Delete(@"C:\Temp\pairedData1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\pairedData1.xlsx");
            de.Read("Sheet1");
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

            ExcelReader de = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(TestUtility.OutputPath + "CommandLineImport1.dss"))
            {
                w.Write(de.Read(0) as TimeSeries);
            }
        }

        [TestMethod]
        public void CommandLineImport2()
        {
            var fn = TestUtility.BasePath + "indexedRegularTimeSeries1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport2.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            ExcelReader de = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(TestUtility.OutputPath + "CommandLineImport2.dss"))
            {
                w.Write(de.Read(0) as TimeSeries);
            }
        }

        [TestMethod]
        public void CommandLineImport3()
        {
            var fn = TestUtility.BasePath + "irregularTimeSeries1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport3.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            ExcelReader de = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(TestUtility.OutputPath + "CommandLineImport3.dss"))
            {
                w.Write(de.Read(0) as TimeSeries);
            }
        }

        [TestMethod]
        public void CommandLineImport4()
        {
            var fn = TestUtility.BasePath + "indexedIrregularTimeSeries1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport4.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            ExcelReader de = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(TestUtility.OutputPath + "CommandLineImport4.dss"))
            {
                w.Write(de.Read(0) as TimeSeries);
            }
        }

        [TestMethod]
        public void CommandLineImport5()
        {
            var fn = TestUtility.BasePath + "pairedData1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport5.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            ExcelReader de = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(TestUtility.OutputPath + "CommandLineImport5.dss"))
            {
                w.Write(de.Read(0) as PairedData);
            }
        }

        [TestMethod]
        public void CommandLineImport6()
        {
            var fn = TestUtility.BasePath + "indexedPairedData1.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport6.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            ExcelReader de = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(TestUtility.OutputPath + "CommandLineImport6.dss"))
            {
                w.Write(de.Read(0) as PairedData);
            }
        }

        [TestMethod]
        public void CommandLineImport7()
        {
            var fn = TestUtility.BasePath + "pairedData2.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport7.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            ExcelReader de = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(TestUtility.OutputPath + "CommandLineImport7.dss"))
            {
                w.Write(de.Read(0) as PairedData);
            }
        }

        [TestMethod]
        public void CommandLineImport8()
        {
            var fn = TestUtility.BasePath + "indexedPairedData2.xlsx";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "CommandLineImport8.dss" };
            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
            }

            ExcelReader de = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(TestUtility.OutputPath + "CommandLineImport8.dss"))
            {
                w.Write(de.Read(0) as PairedData);
            }
        }
    }
}
