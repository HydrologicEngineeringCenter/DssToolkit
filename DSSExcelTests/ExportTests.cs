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

        [TestMethod]
        public void CommandLineExport1()
        {
            var fn = TestUtility.BasePath + "regularTimeSeries1.dss";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "regularTimeSeries1.dss", "/excel/import/plugin/01May2020/15Minute/regularTimeSeriesWLF/"};

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type == RecordType.RegularTimeSeries || type == RecordType.IrregularTimeSeries)
                    record = r.GetTimeSeries(path);
                else if (type == RecordType.PairedData)
                    record = r.GetPairedData(path.FullPath);
                else
                    return;

                DSSExcel.Export(a[2], record);
            }
        }

        [TestMethod]
        public void CommandLineExport2()
        {
            var fn = TestUtility.BasePath + "indexedRegularTimeSeries1.dss";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "indexedRegularTimeSeries1.dss", "/excel/import/plugin/01May2020/15Minute/regularTimeSeriesZ1K/" };

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type == RecordType.RegularTimeSeries || type == RecordType.IrregularTimeSeries)
                    record = r.GetTimeSeries(path);
                else if (type == RecordType.PairedData)
                    record = r.GetPairedData(path.FullPath);
                else
                    return;

                DSSExcel.Export(a[2], record);
            }
        }

        [TestMethod]
        public void CommandLineExport3()
        {
            var fn = TestUtility.BasePath + "irregularTimeSeries1.dss";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "irregularTimeSeries1.dss", "/excel/import/plugin/01Jan2020/IR-Year/irregularTimeSeries52Z/" };

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type == RecordType.RegularTimeSeries || type == RecordType.IrregularTimeSeries)
                    record = r.GetTimeSeries(path);
                else if (type == RecordType.PairedData)
                    record = r.GetPairedData(path.FullPath);
                else
                    return;

                DSSExcel.Export(a[2], record);
            }
        }

        [TestMethod]
        public void CommandLineExport4()
        {
            var fn = TestUtility.BasePath + "indexedIrregularTimeSeries1.dss";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "indexedIrregularTimeSeries1.dss", "/excel/import/plugin/01Jan2020/IR-Year/irregularTimeSeriesM7I/" };

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type == RecordType.RegularTimeSeries || type == RecordType.IrregularTimeSeries)
                    record = r.GetTimeSeries(path);
                else if (type == RecordType.PairedData)
                    record = r.GetPairedData(path.FullPath);
                else
                    return;

                DSSExcel.Export(a[2], record);
            }
        }

        [TestMethod]
        public void CommandLineExport5()
        {
            var fn = TestUtility.BasePath + "pairedData1.dss";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "pairedData1.dss", "/excel/import/plugin//e/pairedDataAXD/" };

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type == RecordType.RegularTimeSeries || type == RecordType.IrregularTimeSeries)
                    record = r.GetTimeSeries(path);
                else if (type == RecordType.PairedData)
                    record = r.GetPairedData(path.FullPath);
                else
                    return;

                DSSExcel.Export(a[2], record);
            }
        }

        [TestMethod]
        public void CommandLineExport6()
        {
            var fn = TestUtility.BasePath + "indexedPairedData1.dss";
            string[] a = new string[] { "import", fn, TestUtility.OutputPath + "indexedPairedData1.dss", "/excel/import/plugin//e/pairedData3CR/" };

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type == RecordType.RegularTimeSeries || type == RecordType.IrregularTimeSeries)
                    record = r.GetTimeSeries(path);
                else if (type == RecordType.PairedData)
                    record = r.GetPairedData(path.FullPath);
                else
                    return;

                DSSExcel.Export(a[2], record);
            }
        }


    }
}
