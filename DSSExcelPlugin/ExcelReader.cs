using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetGear;
using SpreadsheetGear.Advanced.Cells;

namespace Hec.Dss.Excel
{
    public class ExcelReader : ExcelTools
    {

        public ExcelReader(string filename)
        {
            workbook = workbookSet.Workbooks.Open(filename);

        }

        public TimeSeries GetTimeSeries(string worksheet)
        {
            if (!isIrregularTimeSeries(worksheet) && !isRegularTimeSeries(worksheet))
                return new TimeSeries();

            TimeSeries ts = new TimeSeries();
            ts.Times = GetTimeSeriesTimes(worksheet);
            ts.Values = GetTimeSeriesValues(worksheet);
            ts.Path = GetRandomTimeSeriesPath(ts, worksheet);
            ts.DataType = "type1";
            ts.Units = "unit1";

            return ts;
        }

        private DssPath GetRandomTimeSeriesPath(TimeSeries ts, string worksheet)
        {
            if (IsRegular(ts.Times.ToList()))
            {
                var temp = ts;
                temp.Path = new DssPath("import", Path.GetFileNameWithoutExtension(workbook.FullName), worksheet, "", "", "regularTimeSeries" + RandomString(3));
                temp.Path.Epart = TimeWindow.GetInterval(temp);
                return temp.Path;
            }
            else
            {
                return new DssPath("import", Path.GetFileNameWithoutExtension(workbook.FullName), worksheet, "", "IR-Year", "irregularTimeSeries" + RandomString(3));
            }
        }

        public TimeSeries GetTimeSeries(int worksheetIndex)
        {
            return GetTimeSeries(workbook.Worksheets[worksheetIndex].Name);
        }

        private double[] GetTimeSeriesValues(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var r = SmallestColumnRowCount(worksheet);
            var v = new List<double>();
            if (HasIndex(worksheet))
            {
                for (int i = DataStartIndex(worksheet); i < r; i++)
                {
                    v.Add(vals[i, 2].Number);
                }
            }
            else
            {
                for (int i = DataStartIndex(worksheet); i < r; i++)
                {
                    v.Add(vals[i, 1].Number);
                }
            }
            return v.ToArray();
        }

        private DateTime[] GetTimeSeriesTimes(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var r = SmallestColumnRowCount(worksheet);
            var d = new List<DateTime>();
            if (HasIndex(worksheet))
            {
                for (int i = DataStartIndex(worksheet); i < r; i++)
                {
                    d.Add(GetDateFromCell(vals[i, 1].Number));
                }
            }
            else
            {
                for (int i = DataStartIndex(worksheet); i < r; i++)
                {
                    d.Add(GetDateFromCell(vals[i, 0].Number));
                }
            }
            return d.ToArray();
        }

        public PairedData GetPairedData(string worksheet)
        {
            if (!isPairedData(worksheet))
                return new PairedData();

            double[] ordinates = GetPairedDataOrdinates(worksheet);
            List<double[]> vals = GetPairedDataValues(worksheet);
            PairedData pd = new PairedData(ordinates, vals, new List<string>(), "", "", "", "", GetRandomPairedDataPath(worksheet).FullPath);
            pd.UnitsDependent = "unit1";
            pd.UnitsIndependent = "unit2";
            pd.TypeDependent = "type1";
            pd.TypeIndependent = "type2";
            pd.Labels = new List<string>();
            return pd;
        }

        private DssPath GetRandomPairedDataPath(string worksheet)
        {
            return new DssPath("import", Path.GetFileNameWithoutExtension(workbook.FullName), worksheet, "", "excel", "pairedData" + RandomString(3));
        }

        public PairedData GetPairedData(int worksheetIndex)
        {
            return GetPairedData(workbook.Worksheets[worksheetIndex].Name);
        }

        private double[] GetPairedDataOrdinates(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var r = SmallestColumnRowCount(worksheet);
            var o = new List<double>();
            if (HasIndex(worksheet))
            {
                for (int i = DataStartIndex(worksheet); i < r; i++)
                {
                    o.Add(vals[i, 1].Number);
                }
            }
            else
            {
                for (int i = DataStartIndex(worksheet); i < r; i++)
                {
                    o.Add(vals[i, 0].Number);
                }
            }
            return o.ToArray();
        }

        private List<double[]> GetPairedDataValues(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var r = SmallestColumnRowCount(worksheet);
            var c = ColumnCount(worksheet);
            var t = new List<double>();
            var v = new List<double[]>();

            if (HasIndex(worksheet))
            {
                for (int i = 2; i < c; i++)
                {
                    for (int j = DataStartIndex(worksheet); j < r; j++)
                    {
                        t.Add(vals[j, i].Number);
                    }
                    v.Add(t.ToArray());
                    t.Clear();
                }
            }
            else
            {
                for (int i = 1; i < c; i++)
                {
                    for (int j = DataStartIndex(worksheet); j < r; j++)
                    {
                        t.Add(vals[j, i].Number);
                    }
                    v.Add(t.ToArray());
                    t.Clear();
                }
            }
            return v;
        }

        /// <summary>
        /// Returns the DSS data from an excel sheet. 
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public object Read(string sheet)
        {
            var t = CheckType(sheet);
            if (t == RecordType.RegularTimeSeries || t == RecordType.IrregularTimeSeries)
                return GetTimeSeries(sheet);
            else if (t == RecordType.PairedData)
                return GetPairedData(sheet);
            else
                return null;
        }

        public object Read(int sheetIndex)
        {
            return Read(workbook.Worksheets[sheetIndex].Name);
        }

    }
}
