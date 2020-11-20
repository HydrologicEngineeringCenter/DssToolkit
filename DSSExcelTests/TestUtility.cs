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
        public static string SimpleIrregularTSPath = BasePath + "small-ir-ts.xlsx";
        public static string SimpleRegularTSPath = BasePath + "small-r-ts.xlsx";
        public static string SimplePDPath = BasePath + "small-pd.xlsx";

        public static TimeSeries CreateTimeSeries(int numberOfVals, bool regular = true)
        {
            List<double> vals = new List<double>();
            List<DateTime> dt = new List<DateTime>();
            var d = new DateTime(2020, 1, 1);

            if (regular)
            {
                for (int i = 0; i < numberOfVals; i++)
                {
                    vals.Add(i * 2);
                    dt.Add(d.AddDays(i));
                }
                var ts = new TimeSeries(new DssPath("a", "b", "c", "", "1Day", "f"), vals.ToArray(), d, "", "");
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
                var ts = new TimeSeries(new DssPath("a", "b", "c", "", "IR-Year", "f"), vals.ToArray(), d, "", "");
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

        public static List<DateTime> CreateTimeSeriesTimes(string start, int count, double days, double hours, double minutes)
        {
            var expected_times = new List<DateTime>();
            for (int i = 0; i < count; i++)
            {
                if (i == 0)
                    expected_times.Add(DateTime.Parse(start));
                else
                    expected_times.Add(expected_times[i - 1].AddDays(days).AddHours(hours).AddMinutes(minutes));

            }
            return expected_times;
        }
    }
}
