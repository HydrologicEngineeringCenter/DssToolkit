using Hec.Dss;
using Hec.Excel;
using System;
using System.Collections.Generic;
using System.IO;

namespace UnitTests
{
  public class TestUtility
    {
        internal static string BasePath = @"..\..\..\test-files\";

    static string FullPath(string fileName)
    {
      return Path.Combine(BasePath, fileName);
    }

    public static TimeSeries TimeSeriesFromExcel(string fileName)
    {
      var tsList = ExcelTimeSeries.Read(FullPath(fileName));
      return tsList[0];
    }

    public static TimeSeries TimeSeriesFromDss(string fileName, string path)
    {
      using (DssReader r = new DssReader(FullPath( fileName)))
      {
        TimeSeries ts = r.GetTimeSeries(new DssPath(path));
        return ts;
      } 
    }

    /// <summary>
    /// gets a unique filename that can be created
    /// </summary>
    /// <param name="extension">file extension icluding dot</param>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    public static string GetSimpleTempFileName(string extension)
    {
      string path = Path.GetTempPath();

      var part1 = Guid.NewGuid().ToString("N").Substring(0, 8).ToLower();
      var fn = Path.Combine(path, part1 + extension);

      if (File.Exists(fn))
        throw new Exception("File allready exists");

      return fn;
    }


    //static string OutputPath = @"..\..\test-files\output\";
    // static string SimpleIrregularTSPath = BasePath + "small-ir-ts.xlsx";
    //  static string SimpleRegularTSPath = BasePath + "small-r-ts.xlsx";
    // static string SimplePDPath = BasePath + "small-pd.xlsx";

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
            pd.Path = new DssPath("a", "b", "c", "", "e", "f");
            pd.UnitsIndependent = "unit 1";
            pd.UnitsDependent = "unit 2";
            pd.TypeIndependent = "type 1";
            pd.TypeDependent = "type 2";
            return pd;
        }

        private static Random random = new Random();
       

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
