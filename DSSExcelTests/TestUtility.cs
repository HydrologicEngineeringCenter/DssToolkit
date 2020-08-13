using Hec.Dss;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSSExcelTests
{
    public class TestUtility
    {
        public static string BasePath = @"..\..\test-files\";
        public static string OutputPath = @"..\..\test-files\output\";

        public static TimeSeries CreateTimeSeries(int numberOfVals, bool regular = true)
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

        public static PairedData CreatePairedData(int numberOfCurves, int numberOfVals)
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
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
