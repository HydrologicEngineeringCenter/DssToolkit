using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetGear;
using SpreadsheetGear.Advanced.Cells;
using static Hec.Dss.Excel.ExcelTools;

namespace Hec.Dss.Excel
{
    public class ExcelReader
    {
        public IWorkbookSet workbookSet = Factory.GetWorkbookSet();
        public IWorkbook workbook;
        public SheetInfo ActiveSheetInfo { get; private set; }
        public int WorksheetCount
        {
            get
            {
                return workbook.Worksheets.Count;
            }
        }

        /// <summary>
        /// Get sheet info for a specific sheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public SheetInfo GetWorksheetInfo(string worksheet)
        {
            return ActiveSheetInfo != null && ActiveSheetInfo.Name == worksheet ? ActiveSheetInfo : new SheetInfo(this, worksheet);
        }

        public ExcelReader(string filename)
        {
            workbook = workbookSet.Workbooks.Open(filename);
        }

        public ExcelReader(IWorkbook workbook)
        {
            this.workbook = workbook;
        }

        public TimeSeries GetTimeSeries(string worksheet)
        {
            ActiveSheetInfo = GetWorksheetInfo(worksheet);
            if (!isIrregularTimeSeries(worksheet) && !isRegularTimeSeries(worksheet))
                return new TimeSeries();

            TimeSeries ts = new TimeSeries();
            GetTimeSeriesTimes(ts, worksheet);
            GetTimeSeriesValues(ts, worksheet, ActiveSheetInfo.ValueStartColumnIndex);
            GetTimeSeriesPath(ts, worksheet, ActiveSheetInfo.ValueStartColumnIndex);
            GetTimeSeriesUnits(ts, worksheet, ActiveSheetInfo.ValueStartColumnIndex);
            GetTimeSeriesDataType(ts, worksheet, ActiveSheetInfo.ValueStartColumnIndex);

            return ts;
        }

        public IEnumerable<TimeSeries> GetMultipleTimeSeries(string worksheet)
        {
            ActiveSheetInfo = GetWorksheetInfo(worksheet);
            if (!isIrregularTimeSeries(worksheet) && !isRegularTimeSeries(worksheet))
                return new List<TimeSeries>();
            var l = new List<TimeSeries>();
            for (int i = ActiveSheetInfo.ValueStartColumnIndex; i < ActiveSheetInfo.ColumnCount; i++)
            {
                TimeSeries ts = new TimeSeries();
                GetTimeSeriesTimes(ts, worksheet);
                GetTimeSeriesValues(ts, worksheet, i);
                GetTimeSeriesPath(ts, worksheet, i);
                GetTimeSeriesUnits(ts, worksheet, i);
                GetTimeSeriesDataType(ts, worksheet, i);
                l.Add(ts);
            }
            return l;
        }

        private void GetTimeSeriesDataType(TimeSeries ts, string worksheet, int valueColumn)
        {
            var s = "DataType";
            if (ActiveSheetInfo.PathLayout != PathLayout.NoPath &&
                ActiveSheetInfo.PathLayout != PathLayout.TS_PathWithoutTypeAndUnits &&
                ActiveSheetInfo.PathLayout != PathLayout.TS_PathWithoutDPartTypeAndUnit)
            {
                int dataTypeIndex = (int)ActiveSheetInfo.PathLayout - 1;
                s = CellToString(workbook.Worksheets[worksheet].Cells[dataTypeIndex, valueColumn]);
                ts.DataType = s;
            }

        }

        private void GetTimeSeriesUnits(TimeSeries ts, string worksheet, int valueColumn)
        {
            var s = "Unit";
            if (ActiveSheetInfo.PathLayout != PathLayout.NoPath &&
                ActiveSheetInfo.PathLayout != PathLayout.TS_PathWithoutTypeAndUnits &&
                ActiveSheetInfo.PathLayout != PathLayout.TS_PathWithoutDPartTypeAndUnit)
            {
                int unitIndex = (int)ActiveSheetInfo.PathLayout - 2;
                s = CellToString(workbook.Worksheets[worksheet].Cells[unitIndex, valueColumn]);
                ts.Units = s;
            }

        }

        private void GetTimeSeriesPath(TimeSeries ts, string worksheet, int valueColumn)
        {
            if (!ActiveSheetInfo.HasPath)
            {
                GetRandomTimeSeriesPath(ts, worksheet);
                return;
            }

            GetPath(ts, worksheet, valueColumn, ActiveSheetInfo.PathLayout);
            if (!PathPartsAreValid(ts.Path))
            {
                GetRandomTimeSeriesPath(ts, worksheet);
            }
        }

        private void GetPath(TimeSeries ts, string worksheet, int column, PathLayout pathLayout)
        {
            ts.Path = new DssPath();
            if (pathLayout == PathLayout.TS_StandardPath || pathLayout == PathLayout.TS_PathWithoutTypeAndUnits)
            {
                ts.Path.Apart = CellToString(workbook.Worksheets[worksheet].Cells[0, column]);
                ts.Path.Bpart = CellToString(workbook.Worksheets[worksheet].Cells[1, column]);
                ts.Path.Cpart = CellToString(workbook.Worksheets[worksheet].Cells[2, column]);
                ts.Path.Fpart = CellToString(workbook.Worksheets[worksheet].Cells[5, column]);
            }
            else if (pathLayout == PathLayout.TS_PathWithoutDPart || pathLayout == PathLayout.TS_PathWithoutDPartTypeAndUnit)
            {
                ts.Path.Apart = CellToString(workbook.Worksheets[worksheet].Cells[0, column]);
                ts.Path.Bpart = CellToString(workbook.Worksheets[worksheet].Cells[1, column]);
                ts.Path.Cpart = CellToString(workbook.Worksheets[worksheet].Cells[2, column]);
                ts.Path.Fpart = CellToString(workbook.Worksheets[worksheet].Cells[4, column]);
            }

            if (IsRegular(ts.Times.ToList()))
                ts.Path.Epart = TimeWindow.GetInterval(ts);
            else
                ts.Path.Epart = "IR-Year";
        }

        private bool PathPartsAreValid(DssPath path)
        {
            return path.Apart != "" || path.Bpart != "" || path.Cpart != "" || path.Fpart != "";
        }
        /// <summary>
        /// Get all values from a specified value column number in a worksheet (non-zero-based indexing).
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="valueColumn"></param>
        /// <returns></returns>
        private void GetTimeSeriesValues(TimeSeries ts, string worksheet, int valueColumn)
        {
            var vals = Values(worksheet);
            var v = new List<double>();
            for (int i = ActiveSheetInfo.DataStartRowIndex; i < ActiveSheetInfo.SmallestColumnRowCount; i++)
                v.Add(vals[i, valueColumn].Number);
             ts.Values = v.ToArray();
        }

        private int TimeSeriesValueColumnCount()
        {
            return ActiveSheetInfo.HasIndex ? ActiveSheetInfo.ColumnCount - 2 : ActiveSheetInfo.ColumnCount - 1;
        }

        private DssPath GetRandomTimeSeriesPath(TimeSeries ts, string worksheet)
        {
            if (IsRegular(ts.Times.ToList()))
            {
                var temp = ts;
                temp.Path = new DssPath("import", Path.GetFileNameWithoutExtension(workbook.FullName), worksheet, 
                    "", "", "regularTimeSeries" + ExcelTools.RandomString(3));
                temp.Path.Epart = TimeWindow.GetInterval(temp);
                return temp.Path;
            }
            else
            {
                return new DssPath("import", Path.GetFileNameWithoutExtension(workbook.FullName), worksheet, 
                    "", "IR-Year", "irregularTimeSeries" + ExcelTools.RandomString(3));
            }
        }

        public TimeSeries GetTimeSeries(int worksheetIndex)
        {
            return GetTimeSeries(workbook.Worksheets[worksheetIndex].Name);
        }

        private void GetTimeSeriesTimes(TimeSeries ts, string worksheet)
        {
            var d = new List<DateTime>();
            var offset = ActiveSheetInfo.HasIndex ? 1 : 0;
            for (int i = ActiveSheetInfo.DataStartRowIndex; i < ActiveSheetInfo.SmallestColumnRowCount; i++)
                d.Add(GetDateFromCell(CellToString(workbook.Worksheets[worksheet].Cells[i, offset])));
            ts.Times = d.ToArray();
        }

        public PairedData GetPairedData(string worksheet)
        {
            ActiveSheetInfo = GetWorksheetInfo(worksheet);
            if (!isPairedData(worksheet))
                return new PairedData();

            PairedData pd = new PairedData();
            GetPairedDataOrdinates(pd, worksheet);
            GetPairedDataValues(pd, worksheet);
            GetPairedDataPath(pd, worksheet);
            GetPairedDataUnits(pd, worksheet);
            GetPairedDataTypes(pd, worksheet);
            GetPairedDataLabels(pd, worksheet);
            return pd;
        }

        private void GetPairedDataPath(PairedData pd, string worksheet)
        {
            if (!ActiveSheetInfo.HasPath)
            {
                GetRandomPairedDataPath(pd, worksheet);
                return;
            }

            int column = 1;
            GetPath(pd, worksheet, column, ActiveSheetInfo.PathLayout);
            if (!PathPartsAreValid(pd.Path))
            {
                GetRandomPairedDataPath(pd, worksheet);
            }
        }

        private void GetPath(PairedData pd, string worksheet, int column, PathLayout pathLayout)
        {
            pd.Path = new DssPath();
            if (pathLayout == PathLayout.PD_StandardPath || pathLayout == PathLayout.PD_PathWithoutTypesAndUnits ||
                pathLayout == PathLayout.PD_PathWithoutTypes || pathLayout == PathLayout.PD_PathWithoutUnits)
            {
                pd.Path.Apart = CellToString(workbook.Worksheets[worksheet].Cells[0, column]);
                pd.Path.Bpart = CellToString(workbook.Worksheets[worksheet].Cells[1, column]);
                pd.Path.Cpart = CellToString(workbook.Worksheets[worksheet].Cells[2, column]);
                pd.Path.Epart = CellToString(workbook.Worksheets[worksheet].Cells[4, column]);
                pd.Path.Fpart = CellToString(workbook.Worksheets[worksheet].Cells[5, column]);
            }
            else if (pathLayout == PathLayout.PD_PathWithoutDPart || pathLayout == PathLayout.PD_PathWithoutDPartTypesAndUnits ||
                pathLayout == PathLayout.PD_PathWithoutDPartAndTypes || pathLayout == PathLayout.PD_PathWithoutDPartAndUnits)
            {
                pd.Path.Apart = CellToString(workbook.Worksheets[worksheet].Cells[0, column]);
                pd.Path.Bpart = CellToString(workbook.Worksheets[worksheet].Cells[1, column]);
                pd.Path.Cpart = CellToString(workbook.Worksheets[worksheet].Cells[2, column]);
                pd.Path.Epart = CellToString(workbook.Worksheets[worksheet].Cells[3, column]);
                pd.Path.Fpart = CellToString(workbook.Worksheets[worksheet].Cells[4, column]);
            }
        }

        private void GetPairedDataLabels(PairedData pd, string worksheet)
        {
            if (!HasLabels())
                return;

            pd.Labels = new List<string>();
            for (int i = ActiveSheetInfo.ValueStartColumnIndex; i < ActiveSheetInfo.ColumnCount; i++)
            {
                pd.Labels.Add(CellToString(workbook.Worksheets[worksheet].Cells[ActiveSheetInfo.DataStartRowIndex - 1, i]));
            }
        }

        private bool HasLabels()
        {
            return ActiveSheetInfo.DataStartRowIndex != 0 && 
                ActiveSheetInfo.DataStartRowIndex != ActiveSheetInfo.PathEndRowIndex + 1;
        }

        private void GetPairedDataTypes(PairedData pd, string worksheet)
        {
            string typeI = "Independent Type";
            string typeD = "Dependent Type";
            if (HasTypes())
            {
                int adjustment1 = 2;
                int adjustment2 = 1;
                int typeIIndex = (int)ActiveSheetInfo.PathLayout - adjustment1;
                int typeDIndex = (int)ActiveSheetInfo.PathLayout - adjustment2;
                int column = 1;
                typeI = CellToString(workbook.Worksheets[worksheet].Cells[typeIIndex, column]);
                typeD = CellToString(workbook.Worksheets[worksheet].Cells[typeDIndex, column]);
            }
            pd.TypeIndependent = typeI;
            pd.TypeDependent = typeD;
        }

        private void GetPairedDataUnits(PairedData pd, string worksheet)
        {
            string unitI = "Independent Unit";
            string unitD = "Dependent Unit";
            if (HasUnits())
            {
                int adjustment1 = 2;
                int adjustment2 = 1;
                if (HasTypes())
                {
                    adjustment1 = 4;
                    adjustment2 = 3;
                }
                int unitIIndex = (int)ActiveSheetInfo.PathLayout - adjustment1;
                int unitDIndex = (int)ActiveSheetInfo.PathLayout - adjustment2;
                int column = 1;
                unitI = CellToString(workbook.Worksheets[worksheet].Cells[unitIIndex, column]);
                unitD = CellToString(workbook.Worksheets[worksheet].Cells[unitDIndex, column]);
            }
            pd.UnitsIndependent = unitI;
            pd.UnitsDependent = unitD;
        }

        private bool HasTypes()
        {
            return ActiveSheetInfo.PathLayout == PathLayout.PD_StandardPath ||
                ActiveSheetInfo.PathLayout == PathLayout.PD_PathWithoutDPart ||
                ActiveSheetInfo.PathLayout == PathLayout.PD_PathWithoutDPartAndUnits ||
                ActiveSheetInfo.PathLayout == PathLayout.PD_PathWithoutUnits;
        }

        private bool HasUnits()
        {
            return ActiveSheetInfo.PathLayout == PathLayout.PD_StandardPath ||
                ActiveSheetInfo.PathLayout == PathLayout.PD_PathWithoutTypes ||
                ActiveSheetInfo.PathLayout == PathLayout.PD_PathWithoutDPart ||
                ActiveSheetInfo.PathLayout == PathLayout.PD_PathWithoutDPartAndTypes;
        }

        private DssPath GetRandomPairedDataPath(PairedData pd, string worksheet)
        {
            return new DssPath("import", Path.GetFileNameWithoutExtension(workbook.FullName), worksheet, "", "excel", "pairedData" + ExcelTools.RandomString(3));
        }

        public PairedData GetPairedData(int worksheetIndex)
        {
            return GetPairedData(workbook.Worksheets[worksheetIndex].Name);
        }

        private void GetPairedDataOrdinates(PairedData pd, string worksheet)
        {
            var vals = Values(worksheet);
            var temp = new List<double>();
            int column = 0;
            for (int i = ActiveSheetInfo.DataStartRowIndex; i < ActiveSheetInfo.SmallestColumnRowCount; i++)
                temp.Add(vals[i, column].Number);
            pd.Ordinates = temp.ToArray();
        }

        private void GetPairedDataValues(PairedData pd, string worksheet)
        {
            var vals = Values(worksheet);
            var temp = new List<double>();
            pd.Values = new List<double[]>();
            var offset = ActiveSheetInfo.HasIndex ? 2 : 1;
            for (int i = offset; i < ActiveSheetInfo.ColumnCount; i++)
            {
                for (int j = ActiveSheetInfo.DataStartRowIndex; j < ActiveSheetInfo.SmallestColumnRowCount; j++)
                    temp.Add(vals[j, i].Number);
                pd.Values.Add(temp.ToArray());
                temp.Clear();
            }
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

        public int DSSPathEndRowIndex(string worksheet)
        {
            int headerRow = 1;
            int dataStartRow = 1;
            int index = DataStartRowIndex(worksheet);
            return index == 0 ? -1 : index - headerRow - dataStartRow; // remove the header and data start rows from data start index to get path end index
        }

        public int DSSPathEndRow(string worksheet)
        {
            return DSSPathEndRowIndex(worksheet) + 1; // remove the header and data start rows from data start row to get path end row
        }

        public bool DSSPathExists(string worksheet, int column)
        {
            int pathEndRow = DSSPathEndRow(worksheet);
            if (pathEndRow < (int)PathLayout.TS_PathWithoutDPartTypeAndUnit ||
                pathEndRow > (int)PathLayout.PD_StandardPath)
                return false;

            int blankEntries = 0;
            for (int i = 0; i < pathEndRow; i++) // check if all entries are blank
            {
                if (!IsValidCell(workbook.Worksheets[worksheet].Cells[i, column]))
                    blankEntries++;
            }

            return blankEntries < (int)PathLayout.TS_PathWithoutDPartTypeAndUnit; // return true if blank entries is less than the amount of entries for a minimal path
        }

        public PathLayout GetDSSPathLayout(string worksheet)
        {
            return (PathLayout)DSSPathEndRow(worksheet);
        }

        public int DataStartRowIndex(string worksheet)
        {
            var r = RowCount(worksheet);
            var c = ColumnCount(worksheet);
            for (int j = 0; j < c; j++)
            {
                for (int i = 0; i < r; i++)
                {
                    if (IsValue(workbook.Worksheets[worksheet].Cells[i, j]))
                        return i;
                }
            }
            return -1;
        }

        public int DataStartRow(string worksheet)
        {
            return DataStartRowIndex(worksheet) + 1;
        }

        public RecordType CheckType(string worksheet)
        {
            ActiveSheetInfo = GetWorksheetInfo(worksheet);
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

        public bool isRegularTimeSeries(string worksheet)
        {
            if (!HasDate(worksheet))
                return false;

            var d = new List<DateTime>();
            var offset = ActiveSheetInfo.HasIndex ? 1 : 0;
            for (int i = ActiveSheetInfo.DataStartRowIndex; i < ActiveSheetInfo.SmallestColumnRowCount; i++)
                d.Add(GetDateFromCell(CellToString(workbook.Worksheets[worksheet].Cells[i, offset])));
            if (IsRegular(d))
                return true;
            return false;
        }

        public IValues Values(string worksheet)
        {
            return (IValues)workbook.Worksheets[worksheet];
        }

        public bool isIrregularTimeSeries(string worksheet)
        {
            if (!HasDate(worksheet))
                return false;

            return !isRegularTimeSeries(worksheet);
        }

        public bool isPairedData(string worksheet)
        {
            var offset = ActiveSheetInfo.HasIndex ? 1 : 0;
            if (ColumnCount(worksheet) < offset)
                return false;

            for (int i = ActiveSheetInfo.DataStartRowIndex; i < ActiveSheetInfo.SmallestColumnRowCount; i++)
            {
                for (int j = offset; i < ActiveSheetInfo.ColumnCount; i++)
                {
                    if (!IsValue(workbook.Worksheets[worksheet].Cells[i, j]))
                        return false;
                }
            }

            return true;
        }

        public bool isGrid(string worksheet)
        {
            return false;
        }

        public bool isTin(string worksheet)
        {
            return false;
        }

        public bool isLocationInfo(string worksheet)
        {
            return false;
        }

        public bool isText(string worksheet)
        {
            return false;
        }

        public bool HasIndex(string worksheet)
        {
            var vals = Values(worksheet);
            var l = new List<int>();
            var start = DataStartRowIndex(worksheet);
            var end = SmallestColumnRowCount(worksheet);

            for (int i = start; i < end; i++)
                l.Add((int)vals[i, 0].Number);

            return l.ToArray().SequenceEqual(Enumerable.Range(1, l.Count)) ||
                l.ToArray().SequenceEqual(Enumerable.Range(0, l.Count - 1));
        }

        public bool HasDate(string worksheet)
        {
            var cells = (workbook.Worksheets[worksheet]).Cells;
            return HasIndex(worksheet) ? IsDate(cells[SmallestColumnRowCount(worksheet) - 1, 1]) : IsDate(cells[SmallestColumnRowCount(worksheet) - 1, 0]);
        }

        public int RowCount(string worksheet)
        {
            try
            {
                return workbook.Worksheets[worksheet].Cells.CurrentRegion.RowCount;
            }
            catch (NullReferenceException)
            {
                return 0;
            }
        }

        public int ColumnCount(string worksheet)
        {
            try
            {
                return workbook.Worksheets[worksheet].Cells.CurrentRegion.ColumnCount;
            }
            catch (NullReferenceException)
            {
                return 0;
            }
        }

        public DateTime GetDateFromCell(string s)
        {
            CorrectDateFormat(s, out DateTime dt);
            return dt;
        }

        private static bool IsRegular(List<DateTime> times)
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

        public void AddSheet(string sheet)
        {
            var s = workbook.Worksheets.Add();
            s.Name = sheet;
        }

        public void AddSheet(int sheetIndex)
        {
            var s = workbook.Worksheets.Add();
            if (!SheetExists(sheetIndex))
                AddSheet(sheetIndex);
        }

        public bool SheetExists(string sheet)
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (workbook.Worksheets[i].Name.ToLower() == sheet.ToLower())
                    return true;
            }
            return false;
        }

        public bool SheetExists(int sheetIndex)
        {
            return sheetIndex >= 0 && sheetIndex < WorksheetCount;
        }

        public static IEnumerable<TimeSeries> GetTimeSeries(IRange DateTimes, IRange Values, string Apart, string Bpart, string Cpart, string Dpart, string Epart, string Fpart)
        {
            var l = new List<TimeSeries>();
            var c = Values.ColumnCount;
            for (int i = 0; i < c; i++)
            {
                var ts = new TimeSeries();
                ts.Times = RangeToDateTimes(DateTimes);
                ts.Values = RangeToTimeSeriesValues(Values, i);
                if (CheckTimeSeriesType(ts.Times) == RecordType.RegularTimeSeries)
                {
                    ts.Path = new DssPath(Apart, Bpart, Cpart, "", "",
                        "r" + (i + 1).ToString() + Fpart, RecordType.RegularTimeSeries, "type", "units");
                    ts.Path.Epart = TimeWindow.GetInterval(ts);
                }
                else
                    ts.Path = new DssPath(Apart, Bpart, Cpart, "", "IR-Year",
                        "r" + (i + 1).ToString() + Fpart, RecordType.IrregularTimeSeries, "type", "units");
                l.Add(ts);
            }

            return l;
        }

        private static double[] RangeToTimeSeriesValues(IRange values, int columnIndex)
        {
            var d = new List<double>();

            for (int i = 0; i < values.RowCount; i++)
            {
                d.Add(double.Parse(values[i, columnIndex].Value.ToString()));
            }

            return d.ToArray();
        }

        private static DateTime[] RangeToDateTimes(IRange dateTimes)
        {
            var r = new List<DateTime>();
            for (int i = 0; i < dateTimes.RowCount; i++)
            {
                CorrectDateFormat(CellToString(dateTimes[i, 0]), out DateTime tmp);
                r.Add(tmp);
            }
            return r.ToArray();
        }

        private static List<double[]> RangeToPairedDataValues(IRange values)
        {
            var d = new List<List<double>>();

            for (int i = 0; i < values.ColumnCount; i++)
            {
                d.Add(new List<double>());
                for (int j = 0; j < values.RowCount; j++)
                    d[i].Add(double.Parse(values[j, i].Value.ToString()));
            }

            var r = new List<double[]>();
            for (int i = 0; i < d.Count; i++)
            {
                r.Add(d[i].ToArray());
            }

            return r;
        }

        private static double[] RangeToOrdinates(IRange ordinates)
        {
            var d = new List<double>();

            for (int i = 0; i < ordinates.RowCount; i++)
            {
                d.Add(double.Parse(ordinates[i, 0].Value.ToString()));
            }

            return d.ToArray();
        }

        private static RecordType CheckTimeSeriesType(DateTime[] times)
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

        public static bool IsValidCell(IRange cell)
        {
            if (cell[0, 0].Value == null || cell[0, 0].Text.Trim() == "")
                return false;

            return true;
        }

        private static void CorrectDateFormat(string s, out DateTime d)
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

        private static bool IsDifferentDateFromat(string s, out DateTime d)
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
                if (!IsValue(range[i, 0]))
                    return false;
            }
            return true;
        }

        public static bool IsValuesRange(IRange range)
        {
            if (!IsAllColumnRowCountsEqual(range))
                return false;

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

            return double.TryParse(value[0, 0].Text, out _);
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

        public bool IsAllColumnRowCountsEqual(string worksheet)
        {
            ActiveSheetInfo = GetWorksheetInfo(worksheet);
            for (int i = 0; i < ActiveSheetInfo.ColumnCount; i++)
            {
                for (int j = ActiveSheetInfo.DataStartRowIndex; j < ActiveSheetInfo.RowCount; j++)
                {
                    if (!IsDate(workbook.Worksheets[worksheet].Cells[j, i]) && !IsValue(workbook.Worksheets[worksheet].Cells[j, i]) && j != ActiveSheetInfo.RowCount)
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
            int r = RowCount(worksheet);
            int s = r - 1;
            int c = ColumnCount(worksheet);
            int start = DataStartRowIndex(worksheet);
            IRange cells = workbook.Worksheets[worksheet].Cells;
            for (int i = 0; i < c; i++)
            {
                for (int j = r - 1; j > start; j--)
                {
                    if (IsValidCell(cells[j, i]) && (IsDate(cells[j, i]) || IsValue(cells[j, i])))
                    {
                        if (s > j)
                            s = j;
                        break;
                    }
                }
            }
            return s + 1;
        }

        public bool AllPathsAreProper(string worksheet)
        {
            var type = CheckType(worksheet);
            if (type == RecordType.IrregularTimeSeries || type == RecordType.RegularTimeSeries)
            {
                for (int i = ActiveSheetInfo.ValueStartColumnIndex; i < ActiveSheetInfo.ColumnCount; i++)
                {
                    if (!DSSPathExists(worksheet, i))
                        return false;
                }
            }
            else if (type == RecordType.PairedData)
            {
                if (!DSSPathExists(worksheet, ActiveSheetInfo.ValueStartColumnIndex))
                    return false;
            }
            return true;
        }
    }

    public enum PathLayout
    {
        TS_StandardPath = 8,
        TS_PathWithoutDPart = 7,
        TS_PathWithoutTypeAndUnits = 6,
        TS_PathWithoutDPartTypeAndUnit = 5,
        PD_StandardPath = 10,
        PD_PathWithoutDPart = 9,
        PD_PathWithoutTypes = 8,
        PD_PathWithoutUnits = 8,
        PD_PathWithoutDPartAndTypes = 7,
        PD_PathWithoutDPartAndUnits = 7,
        PD_PathWithoutDPartTypesAndUnits = 5,
        PD_PathWithoutTypesAndUnits = 6,
        NoPath = 0
    }

    public enum UnitsAndTypes
    {
        TS_UnitAndType,
        TS_Unit,
        TS_Type,
        PD_UnitsAndTypes,
        PD_Units,
        PD_Types
    }
}
