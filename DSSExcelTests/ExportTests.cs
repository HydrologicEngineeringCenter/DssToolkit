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
        [TestMethod]
        public void ExportRegularTimeSeries()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            var fn = @"C:\Temp\exportRTS1.xls";
            File.Delete(fn);
            TimeSeries ts = CreateTimeSeries(10);
            DSSExcel.Export(fn, ts);
        }

        [TestMethod]
        public void ExportIrregularTimeSeries()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            var fn = @"C:\Temp\exportITS1.xls";
            File.Delete(fn);
            TimeSeries ts = CreateTimeSeries(10, false);
            DSSExcel.Export(fn, ts);

        }

        [TestMethod]
        public void ExportPairedData()
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicenseForTesting();
            var fn = @"C:\Temp\exportPD1.xls";
            File.Delete(fn);
            PairedData pd = CreatePairedData(5, 10);
            DSSExcel.Export(fn, pd);
        }

        private TimeSeries CreateTimeSeries(int numberOfVals, bool regular = true)
        {
            List<double> vals = new List<double>();
            List<DateTime> dt = new List<DateTime>();
            var d = new DateTime(2020, 1, 1);

            if (regular)
            {
                for (int i = 0; i < numberOfVals; i++)
                {
                    vals.Add(i * 10);
                    dt.Add(d.AddDays(i));
                }
                var ts = new TimeSeries(new DssPath(RandomString(2), RandomString(2), RandomString(2), "", "1Day", RandomString(2)), vals.ToArray(), d, "", "");
                ts.Times = dt.ToArray();
                return ts;
            }
            else
            {
                for (int i = 0; i < numberOfVals; i++)
                {
                    vals.Add(i + 1);
                    dt.Add(d.AddDays(i * 2));
                }
                var ts = new TimeSeries(new DssPath(RandomString(2), RandomString(2), RandomString(2), "", "IR-Year", RandomString(2)), vals.ToArray(), d, "", "");
                ts.Times = dt.ToArray();
                return ts;
            }
        }

        private PairedData CreatePairedData(int numberOfCurves, int numberOfVals)
        {
            List<double> ordinates = new List<double>();
            List<List<double>> temp = new List<List<double>>();
            List<double[]> vals = new List<double[]>();
            for (int i = 0; i < numberOfCurves; i++)
            {
                temp.Add(new List<double>());
            }

            for (int i = 0; i < numberOfVals; i++)
            {
                ordinates.Add(i + 3);
            }

            for (int i = 0; i < numberOfCurves; i++)
            {
                for (int j = 0; j < numberOfVals; j++)
                {
                    temp[i].Add(j * i);
                }
            }

            for (int i = 0; i < numberOfCurves; i++)
            {
                vals.Add(temp[i].ToArray());
            }

            var pd = new PairedData(ordinates.ToArray(), vals);
            return pd;
        }

        private static Random random = new Random();
        private string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
