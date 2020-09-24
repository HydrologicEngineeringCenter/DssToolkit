using SpreadsheetGear;
using SpreadsheetGear.Advanced.Cells;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hec.Dss.Excel
{
    public abstract class ExcelTools
    {
        public SpreadsheetGear.IWorkbookSet workbookSet = Factory.GetWorkbookSet();
        public SpreadsheetGear.IWorkbook workbook;

        public int Count 
        { 
            get
            {
                return workbook.Worksheets.Count;
            }
        }

        public IRange Cells
        {
            get
            {
                return workbook.ActiveWorksheet.Cells;
            }
        }

        /// <summary>
        /// Returns the row index where the headers end and the data begins.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        protected int DataStartIndex(string worksheet)
        {
            IValues vals = (IValues)workbook.Worksheets[worksheet];
            var r = RowCount(worksheet);
            var c = ColumnCount(worksheet);
            for (int j = 0; j < c; j++)
            {
                for (int i = 0; i < r; i++)
                {
                    if (vals[i, j].Type == SpreadsheetGear.Advanced.Cells.ValueType.Number)
                        return i;
                }
            }
            return -1;
        }

        public RecordType CheckType(string worksheet)
        {
            if (isRegularTimeSeries(worksheet))
                return RecordType.RegularTimeSeries;
            else if (isIrregularTimeSeries(worksheet))
                return RecordType.IrregularTimeSeries;
            else if (isPairedData(worksheet))
                return RecordType.PairedData;
            else if (isGrid(worksheet))
                return RecordType.Grid;
            else if (isTin(worksheet))
                return RecordType.Tin;
            else if (isLocationInfo(worksheet))
                return RecordType.LocationInfo;
            else if (isText(worksheet))
                return RecordType.Text;
            else
                return RecordType.Unknown;
        }

        public RecordType CheckType(int worksheetIndex)
        {
            return CheckType(workbook.Worksheets[worksheetIndex].Name);
        }

        protected bool isRegularTimeSeries(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var d = new List<DateTime>();
            if (HasDate(worksheet))
            {
                if (HasIndex(worksheet))
                {
                    for (int i = DataStartIndex(worksheet); i < RowCount(worksheet); i++)
                    {

                        DateTime dt = GetDateFromCell(vals[i, 1].Number);
                        d.Add(dt);
                    }
                    if (IsRegular(d))
                        return true;
                    return false;
                }
                else
                {
                    for (int i = DataStartIndex(worksheet); i < RowCount(worksheet); i++)
                    {

                        DateTime dt = GetDateFromCell(vals[i, 0].Number);
                        d.Add(dt);
                    }
                    if (IsRegular(d))
                        return true;
                    return false;
                }
            }
            return false;
        }

        protected bool isIrregularTimeSeries(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var d = new List<DateTime>();
            if (HasDate(worksheet))
            {
                if (HasIndex(worksheet))
                {
                    for (int i = DataStartIndex(worksheet); i < RowCount(worksheet); i++)
                    {
                        DateTime dt = GetDateFromCell(vals[i, 1].Number);
                        d.Add(dt);
                    }
                    if (IsRegular(d))
                        return false;
                    return true;
                }
                else
                {
                    for (int i = DataStartIndex(worksheet); i < RowCount(worksheet); i++)
                    {
                        DateTime dt = GetDateFromCell(vals[i, 0].Number);
                        d.Add(dt);
                    }
                    if (IsRegular(d))
                        return false;
                    return true;
                }
            }
            return false;
        }

        protected bool isPairedData(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var r = SmallestColumnRowCount(worksheet);
            var c = ColumnCount(worksheet);

            if (HasIndex(worksheet))
            {
                if (ColumnCount(worksheet) < 3)
                    return false;
            }
            else
            {
                if (ColumnCount(worksheet) < 2)
                    return false;
            }

            for (int i = DataStartIndex(worksheet); i < r; i++)
            {
                for (int j = 0; i < c; i++)
                {
                    if (vals[i, j].Type != SpreadsheetGear.Advanced.Cells.ValueType.Number)
                        return false;
                }
            }

            return true;

        }

        protected bool isGrid(string worksheet)
        {
            throw new NotImplementedException();
        }

        protected bool isTin(string worksheet)
        {
            throw new NotImplementedException();
        }

        protected bool isLocationInfo(string worksheet)
        {
            throw new NotImplementedException();
        }

        protected bool isText(string worksheet)
        {
            throw new NotImplementedException();
        }

        protected bool HasIndex(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var l = new List<int>();
            var start = DataStartIndex(worksheet);

            if (vals[start, 0].Type != SpreadsheetGear.Advanced.Cells.ValueType.Number &&
                (vals[start, 0].Number != 0 && vals[start, 0].Number != 1))
                return false;

            for (int i = start; i < SmallestColumnRowCount(worksheet); i++)
            {
                l.Add((int)(vals[i, 0].Number));
            }

            return l.ToArray().SequenceEqual(Enumerable.Range(1, l.Count)) ||
                l.ToArray().SequenceEqual(Enumerable.Range(0, l.Count - 1)) ? true : false;

        }

        protected bool HasDate(string worksheet)
        {
            var cells = (workbook.Worksheets[worksheet]).Cells;
            if (HasIndex(worksheet))
            {
                return cells[SmallestColumnRowCount(worksheet) - 1, 1].NumberFormatType == NumberFormatType.DateTime ||
                    cells[SmallestColumnRowCount(worksheet) - 1, 1].NumberFormatType == NumberFormatType.Date;
            }
            else
            {
                return cells[SmallestColumnRowCount(worksheet) - 1, 0].NumberFormatType == NumberFormatType.DateTime ||
                    cells[SmallestColumnRowCount(worksheet) - 1, 0].NumberFormatType == NumberFormatType.Date;
            }
        }

        /// <summary>
        /// Returns default row count of a given worksheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public int RowCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.RowCount;
        }

        public int ColumnCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.ColumnCount;
        }

        /// <summary>
        /// Gets DateTime object from the double value of a date from an excel sheet.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static DateTime GetDateFromCell(double value)
        {
            IWorkbookSet wbs = SpreadsheetGear.Factory.GetWorkbookSet();
            IWorkbook wb = wbs.Workbooks.Add();
            DateTime dt;
            var b = DateTime.TryParse(wb.NumberToDateTime(value).ToString(), out dt);
            return b ? dt : new DateTime();
        }

        public static DateTime GetDateFromCell(string s)
        {
            CorrectDateFormat(s, out DateTime dt);
            return dt;
        }

        public static DateTime[] GetTimeSeriesTimes(IRange range)
        {
            IWorkbookSet wbs = SpreadsheetGear.Factory.GetWorkbookSet();
            IWorkbook wb = wbs.Workbooks.Add();
            var vals = (IValues)range;
            var r = range.RowCount;
            var d = new List<DateTime>();
            for (int i = 0; i < r; i++)
            {
                DateTime dt;
                var b = DateTime.TryParse(wb.NumberToDateTime(vals[i, 0].Number).ToString(), out dt);
                d.Add(b ? dt : new DateTime());
            }
            return d.ToArray();
        }

        public static bool IsRegular(List<DateTime> times)
        {
            var temp = times;
            var td = temp[1] - temp[0];
            for (int i = 0; i < temp.Count; i++)
            {
                if (i == 0)
                    continue;
                else if (i == temp.Count - 1)
                    break;
                else
                {
                    if (temp[i + 1] - temp[i] == td) // check if time difference is the same throughout list
                        continue;
                    else
                        return false;
                }
            }
            return true;
        }

        protected static Random random = new Random();
        public static string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }

        protected void AddSheet(string sheet)
        {
            var s = workbook.Worksheets.Add();
            s.Name = sheet;
        }

        protected void AddSheet(int sheetIndex)
        {
            var s = workbook.Worksheets.Add();
            if (!SheetExists(sheetIndex))
                AddSheet(sheetIndex);
        }

        protected bool SheetExists(string sheet)
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (workbook.Worksheets[i].Name.ToLower() == sheet.ToLower())
                    return true;
            }
            return false;
        }

        protected bool SheetExists(int sheetIndex)
        {
            return sheetIndex >= 0 && sheetIndex < Count;
        }

        protected int IndexOfSheet(string sheet)
        {
            for (int i = 0; i < Count; i++)
            {
                if (workbook.Worksheets[i].Name == sheet)
                    return i;
            }
            return -1;
        }

        public static TimeSeries GetTimeSeries(IRange DateTimes, IRange Values, string Apart, string Bpart, string Cpart, string Dpart, string Epart, string Fpart)
        {
            var ts = new TimeSeries();
            ts.Times = RangeToDateTimes(DateTimes);
            ts.Values = RangeToTimeSeriesValues(Values);
            if (CheckTimeSeriesType(ts.Times) == RecordType.RegularTimeSeries)
            {
                ts.Path = new DssPath(Apart, Bpart, Cpart, "", "", Fpart, RecordType.RegularTimeSeries, "type", "units");
                ts.Path.Epart = TimeWindow.GetInterval(ts);
            }
            else
                ts.Path = new DssPath(Apart, Bpart, Cpart, "", "IR-Year", Fpart, RecordType.IrregularTimeSeries, "type", "units");
            return ts;
        }

        private static double[] RangeToTimeSeriesValues(IRange values)
        {
            var d = new List<double>();

            for (int i = 0; i < values.RowCount; i++)
            {
                d.Add(double.Parse(values[i, 0].Value.ToString()));
            }

            return d.ToArray();
        }

        public static DateTime[] RangeToDateTimes(IRange dateTimes)
        {
            var r = new List<DateTime>();
            for (int i = 0; i < dateTimes.RowCount; i++)
            {
                DateTime tmp;
                CorrectDateFormat(CellToString(dateTimes[i, 0]), out tmp);
                r.Add(tmp);
            }
            return r.ToArray();
        }

        public static PairedData GetPairedData(IRange Ordinates, IRange Values, string Apart, string Bpart, string Cpart, string Dpart, string Epart, string Fpart)
        {
            var pd = new PairedData();
            pd.Ordinates = RangeToOrdinates(Ordinates);
            pd.Values = RangeToPairedDataValues(Values);
            var p = new DssPath(Apart, Bpart, Cpart, Dpart, Epart, Fpart, RecordType.PairedData, "type", "units");
            pd.Path = p;
            return pd;
        }

        public static List<double[]> RangeToPairedDataValues(IRange values)
        {
            var d = new List<List<double>>();

            for (int i = 0; i < values.ColumnCount; i++)
            {
                d.Add(new List<double>());
                for (int j = 0; j < values.RowCount; j++)
                {
                    d[i].Add(double.Parse(values[j, i].Value.ToString()));
                }
            }

            var r = new List<double[]>();
            for (int i = 0; i < d.Count; i++)
            {
                r.Add(d[i].ToArray());
            }

            return r;
        }

        public static double[] RangeToOrdinates(IRange ordinates)
        {
            var d = new List<double>();

            for (int i = 0; i < ordinates.RowCount; i++)
            {
                d.Add(double.Parse(ordinates[i, 0].Value.ToString()));
            }

            return d.ToArray();
        }

        public static RecordType CheckTimeSeriesType(DateTime[] times)
        {
            return IsRegular(times.ToList()) ? RecordType.RegularTimeSeries : RecordType.IrregularTimeSeries;
        }

        public static bool IsDateRange(IRange range)
        {
            for (int i = 0; i < range.RowCount; i++)
            {
                if (!IsDate(range[i, 0]))
                    return false;
            }
            return true;
        }

        public static bool IsDate(IRange date)
        {
            if (!IsValidCell(date))
                return false;

            CorrectDateFormat(date[0, 0].Text, out DateTime d);
            return d == new DateTime() ? false : DateTime.TryParse(d.ToString(), out _);
        }

        public static bool IsValidCell(IRange date)
        {
            if (date[0, 0].Value == null || date[0, 0].Text.Trim() == "")
                return false;

            return true;
        }

        public static void CorrectDateFormat(string s, out DateTime d)
        {
            if (s.Contains("2400") || s.Contains("24:00") || s.Contains("24:00:00"))
            {
                string tmp;
                tmp = s.Replace("2400", "0000");
                tmp = tmp.Replace("24:00", "00:00");
                tmp = tmp.Replace("24:00:00", "00:00:00");
                if (!DateTime.TryParse(tmp, out d))
                    IsDifferentDateFromat(tmp, out d);
                d = d.AddDays(1);
            }
            else
            {
                if (!DateTime.TryParse(s, out d))
                    IsDifferentDateFromat(s, out d);
            }
        }

        public static bool IsDifferentDateFromat(string s, out DateTime d)
        {
            string[] formats =
            {
                "ddMMMyyyy HHmm",
                "ddMMMyyyy HH:mm",
                "ddMMMyyyy HH:mm:ss",
                "ddMMMyyyy  HHmm",
                "ddMMMyyyy  HH:mm",
                "ddMMMyyyy  HH:mm:ss"

            };

            if (DateTime.TryParseExact(s, formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out d))
                return true;
            
            return false;
        }

        public static bool IsOrdinateRange(IRange range)
        {
            return IsValueRange(range);
        }

        public static bool IsValueRange(IRange range)
        {
            for (int i = 0; i < range.RowCount; i++)
            {
                if (!IsValue(range[i,0]))
                    return false;
            }
            return true;
        }

        public static bool IsValuesRange(IRange range)
        {
            for (int i = 0; i < range.RowCount; i++)
            {
                for (int j = 0; j < range.ColumnCount; j++)
                {
                    if (!IsValue(range[i, j]))
                        return false;
                }
            }
            return true;
        }

        public static bool IsValue(IRange value)
        {
            if (!IsValidCell(value))
                return false;

            return double.TryParse(value.Cells[0, 0].Text, out _);
        }

        public static string CellToString(IRange value)
        {
            return value[0, 0].Text;
        }

        public static bool IsAllColumnRowCountsEqual(IRange range)
        {

            for (int i = 0; i < range.ColumnCount; i++)
            {
                for (int j = 0; j < range.RowCount; j++)
                {
                    if (!IsDate(range[j, i]) && !IsValue(range[j, i]) && j != range.RowCount)
                        return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Returns the smallest row count of all columns in a given worksheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public int SmallestColumnRowCount(string worksheet)
        {
            int s = -1;
            for (int i = 0; i < ColumnCount(worksheet); i++)
            {
                for (int j = 0; j < RowCount(worksheet); j++)
                {
                    if (!IsDate(Cells[j, i]) && !IsValue(Cells[j, i]) && j != Cells.RowCount)
                    {
                        if (s == -1 || s > j)
                            s = j;

                    }
                }
            }
            return s;
        }
    }
}