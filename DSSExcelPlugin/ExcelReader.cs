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

        public IEnumerable<TimeSeries> GetMultipleTimeSeries(string worksheet)
        {
            if (!isIrregularTimeSeries(worksheet) && !isRegularTimeSeries(worksheet))
                return new List<TimeSeries>();
            var l = new List<TimeSeries>();
            var c = TimeSeriesValueColumnCount(worksheet);
            for (int i = 0; i < c; i++)
            {
                TimeSeries ts = new TimeSeries();
                ts.Times = GetTimeSeriesTimes(worksheet);
                ts.Values = GetTimeSeriesValues(worksheet, i);
                ts.Path = GetRandomTimeSeriesPath(ts, worksheet);
                ts.DataType = "type1";
                ts.Units = "unit1";
                l.Add(ts);
            }
            return l;
        }

        /// <summary>
        /// Get all values from a specified value column number in a worksheet (non-zero-based indexing).
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="column"></param>
        /// <returns></returns>
        private double[] GetTimeSeriesValues(string worksheet, int column)
        {
            var vals = GetValues(worksheet);
            var r = SmallestColumnRowCount(worksheet);
            var v = new List<double>();
            var start = DataStartIndex(worksheet);
            int offset = HasIndex(worksheet) ? column + 1 : column;
            for (int i = start; i < r; i++)
                v.Add(vals[i, offset].Number);
            return v.ToArray();
        }

        private int TimeSeriesValueColumnCount(string worksheet)
        {
            return HasIndex(worksheet) ? ColumnCount(worksheet) - 2 : ColumnCount(worksheet) - 1;
        }

        private DssPath GetRandomTimeSeriesPath(TimeSeries ts, string worksheet)
        {
            if (IsRegular(ts.Times.ToList()))
            {
                var temp = ts;
                temp.Path = new DssPath("import", Path.GetFileNameWithoutExtension(workbook.FullName), worksheet, 
                    "", "", "regularTimeSeries" + RandomString(3));
                temp.Path.Epart = TimeWindow.GetInterval(temp);
                return temp.Path;
            }
            else
            {
                return new DssPath("import", Path.GetFileNameWithoutExtension(workbook.FullName), worksheet, 
                    "", "IR-Year", "irregularTimeSeries" + RandomString(3));
            }
        }

        public TimeSeries GetTimeSeries(int worksheetIndex)
        {
            return GetTimeSeries(workbook.Worksheets[worksheetIndex].Name);
        }

        /// <summary>
        /// Get all values from the first value column in a worksheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private double[] GetTimeSeriesValues(string worksheet)
        {
            var vals = GetValues(worksheet);
            var r = SmallestColumnRowCount(worksheet);
            var v = new List<double>();
            var start = DataStartIndex(worksheet);
            int offset = HasIndex(worksheet) ? 2 : 1;
            for (int i = start; i < r; i++)
                v.Add(vals[i, offset].Number);
            return v.ToArray();
        }

        private DateTime[] GetTimeSeriesTimes(string worksheet)
        {
            var r = SmallestColumnRowCount(worksheet);
            var d = new List<DateTime>();
            var start = DataStartIndex(worksheet);
            var offset = HasIndex(worksheet) ? 1 : 0;
            for (int i = start; i < r; i++)
                d.Add(GetDateFromCell(CellToString(workbook.Worksheets[worksheet].Cells[i, offset])));
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
            return GetTimeSeriesValues(worksheet);
        }

        private List<double[]> GetPairedDataValues(string worksheet)
        {
            var vals = GetValues(worksheet);
            var r = SmallestColumnRowCount(worksheet);
            var c = ColumnCount(worksheet);
            var t = new List<double>();
            var v = new List<double[]>();
            var start = DataStartIndex(worksheet);
            var offset = HasIndex(worksheet) ? 2 : 1;
            for (int i = offset; i < c; i++)
            {
                for (int j = start; j < r; j++)
                    t.Add(vals[j, i].Number);
                v.Add(t.ToArray());
                t.Clear();
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

        /// <summary>
        /// Read all records that exist in a given sheet.
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        public IEnumerable<object> ReadAll(string sheet)
        {
            var t = CheckType(sheet);
            if (t == RecordType.RegularTimeSeries || t == RecordType.IrregularTimeSeries)
                return GetMultipleTimeSeries(sheet);
            else
                return null;
        }

        public object Read(int sheetIndex)
        {
            return Read(workbook.Worksheets[sheetIndex].Name);
        }

    }
}
