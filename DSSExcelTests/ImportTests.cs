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
        [TestMethod]
        public void ImportRegularTimeSeries1()
        {
            File.Delete(TestUtility.BasePath + "indexedRegularTimeSeries1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\indexedRegularTimeSeries1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportIrregularTimeSeries1()
        {
            File.Delete(TestUtility.BasePath + "indexedIrregularTimeSeries1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\indexedIrregularTimeSeries1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportPairedData1()
        {
            File.Delete(TestUtility.BasePath + "indexedPairedData1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\indexedPairedData1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportRegularTimeSeries2()
        {
            File.Delete(TestUtility.BasePath + "regularTimeSeries1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\regularTimeSeries1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportIrregularTimeSeries2()
        {
            File.Delete(TestUtility.BasePath + "irregularTimeSeries1.dss");
            ExcelReader de = new ExcelReader(@"C:\Temp\irregularTimeSeries1.xlsx");
            de.Read("Sheet1");

        }

        [TestMethod]
        public void ImportPairedData2()
        {
            File.Delete(TestUtility.BasePath + "pairedData1.dss");
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

            ExcelReader er = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(a[2]))
            {
                var t = er.CheckType("sheet1");
                if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                    w.Write(er.Read("sheet1") as TimeSeries);
                else if (t is RecordType.PairedData)
                    w.Write(er.Read("sheet1") as PairedData);

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

            ExcelReader er = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(a[2]))
            {
                var t = er.CheckType("sheet1");
                if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                    w.Write(er.Read("sheet1") as TimeSeries);
                else if (t is RecordType.PairedData)
                    w.Write(er.Read("sheet1") as PairedData);

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

            ExcelReader er = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(a[2]))
            {
                var t = er.CheckType("sheet1");
                if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                    w.Write(er.Read("sheet1") as TimeSeries);
                else if (t is RecordType.PairedData)
                    w.Write(er.Read("sheet1") as PairedData);

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

            ExcelReader er = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(a[2]))
            {
                var t = er.CheckType("sheet1");
                if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                    w.Write(er.Read("sheet1") as TimeSeries);
                else if (t is RecordType.PairedData)
                    w.Write(er.Read("sheet1") as PairedData);

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

            ExcelReader er = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(a[2]))
            {
                var t = er.CheckType("sheet1");
                if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                    w.Write(er.Read("sheet1") as TimeSeries);
                else if (t is RecordType.PairedData)
                    w.Write(er.Read("sheet1") as PairedData);

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

            ExcelReader er = new ExcelReader(a[1]);
            using (DssWriter w = new DssWriter(a[2]))
            {
                var t = er.CheckType("sheet1");
                if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                    w.Write(er.Read("sheet1") as TimeSeries);
                else if (t is RecordType.PairedData)
                    w.Write(er.Read("sheet1") as PairedData);

            }
        }

    }
}
