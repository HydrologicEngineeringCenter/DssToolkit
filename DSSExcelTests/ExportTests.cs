using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Hec.Dss.Excel;
using Hec.Dss;
using System.Data.SqlTypes;
using System.Linq;
using System.IO;
using System.Runtime.InteropServices;

namespace DSSExcelTests
{
    /// <summary>
    /// Summary description for ExportTests
    /// </summary>
    [TestClass]
    public class ExportTests
    {
        [TestMethod]
        public void ExportRegularTimeSeries()
        {
            var fn = TestUtility.BasePath + "exportRTS1.xls";
            File.Delete(fn);
            TimeSeries ts = TestUtility.CreateTimeSeries(10);
            ExcelWriter ew = new ExcelWriter(fn);
            ew.Write(ts, 0);
        }

        [TestMethod]
        public void ExportIrregularTimeSeries()
        {
            var fn = TestUtility.BasePath + "exportITS1.xls";
            File.Delete(fn);
            TimeSeries ts = TestUtility.CreateTimeSeries(10, false);
            ExcelWriter ew = new ExcelWriter(fn);
            ew.Write(ts, 0);

        }

        [TestMethod]
        public void ExportPairedData()
        {
            var fn = TestUtility.BasePath + "exportPD1.xls";
            File.Delete(fn);
            PairedData pd = TestUtility.CreatePairedData(5, 10);
            ExcelWriter ew = new ExcelWriter(fn);
            ew.Write(pd, 0);
        }

        [TestMethod]
        public void CommandLineExport1()
        {
            var fn = TestUtility.BasePath + "regularTimeSeries1.dss";
            string[] a = new string[] { "export", fn, TestUtility.OutputPath + "regularTimeSeries1.xlsx", "/excel/import/plugin//15Minute/regularTimeSeriesWLF/" };

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                ExcelWriter ew = new ExcelWriter(a[2]);
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                {
                    record = r.GetTimeSeries(path);
                    ew.Write(record as TimeSeries, "sheet1");
                }
                else if (type is RecordType.PairedData)
                {
                    record = r.GetPairedData(path.FullPath);
                    ew.Write(record as PairedData, "sheet1");
                }
            }
        }

        [TestMethod]
        public void CommandLineExport2()
        {
            var fn = TestUtility.BasePath + "indexedRegularTimeSeries1.dss";
            string[] a = new string[] { "export", fn, TestUtility.OutputPath + "indexedRegularTimeSeries1.xlsx", "/excel/import/plugin//15Minute/regularTimeSeriesZ1K/" };

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                ExcelWriter ew = new ExcelWriter(a[2]);
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                {
                    record = r.GetTimeSeries(path);
                    ew.Write(record as TimeSeries, "sheet1");
                }
                else if (type is RecordType.PairedData)
                {
                    record = r.GetPairedData(path.FullPath);
                    ew.Write(record as PairedData, "sheet1");
                }
            }
        }

        [TestMethod]
        public void CommandLineExport3()
        {
            var fn = TestUtility.BasePath + "irregularTimeSeries1.dss";
            string[] a = new string[] { "export", fn, TestUtility.OutputPath + "irregularTimeSeries1.xlsx", "/excel/import/plugin/01Jan2020/IR-Year/irregularTimeSeries52Z/" };

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                ExcelWriter ew = new ExcelWriter(a[2]);
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                {
                    record = r.GetTimeSeries(path);
                    ew.Write(record as TimeSeries, "sheet1");
                }
                else if (type is RecordType.PairedData)
                {
                    record = r.GetPairedData(path.FullPath);
                    ew.Write(record as PairedData, "sheet1");
                }
            }
        }

        [TestMethod]
        public void CommandLineExport4()
        {
            var fn = TestUtility.BasePath + "indexedIrregularTimeSeries1.dss";
            string[] a = new string[] { "export", fn, TestUtility.OutputPath + "indexedIrregularTimeSeries1.xlsx", "/excel/import/plugin/01Jan2020/IR-Year/irregularTimeSeriesM7I/" };

            if (!File.Exists(a[1]))
            {
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
            }

            using (DssReader r = new DssReader(a[1]))
            {
                object record;
                ExcelWriter ew = new ExcelWriter(a[2]);
                DssPath path = new DssPath(a[3]);
                var type = r.GetRecordType(path);
                if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                {
                    record = r.GetTimeSeries(path);
                    ew.Write(record as TimeSeries, "sheet1");
                }
                else if (type is RecordType.PairedData)
                {
                    record = r.GetPairedData(path.FullPath);
                    ew.Write(record as PairedData, "sheet1");
                }
            }
        }
        
    }
}
