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
            var fn = TestUtility.BasePath + "l.dss";
            string[] a = new string[] { "export", fn, TestUtility.OutputPath + "l.xlsx", "/Elmore/Atlanta Summit/PRECIP-CUM//IR-Year/SNOTEL/" };

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
            var fn = TestUtility.BasePath + "snotel_example.dss";
            string[] a = new string[] { "export", fn, TestUtility.OutputPath + "snotel_example.xlsx", "/ELMORE/ATLANTA SUMMIT/PRECIP-CUM//IR-MONTH/SNOTEL/" };

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
            var fn = TestUtility.BasePath + "stevenson_elev.dss";
            string[] a = new string[] { "export", fn, TestUtility.OutputPath + "stevenson_elev.xlsx", "//Stevenson/Elev//15Minute/DCP-raw/" };

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
            var fn = TestUtility.BasePath + "WolfCreek_USGS_FlowMeasurements.dss";
            string[] a = new string[] { "export", fn, TestUtility.OutputPath + "WolfCreek_USGS_FlowMeasurements.xlsx", "/CARUTHERS C/IVANPAH CA/FLOW//15Minute/USGS/" };

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
