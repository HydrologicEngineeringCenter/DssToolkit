using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetGear;
using SpreadsheetGear.Advanced.Cells;
using SpreadsheetGear.Shapes;

namespace Hec.Dss.Excel
{
    public class ExcelWriter : IExcelTools
    {
        public IWorkbookSet workbookSet = Factory.GetWorkbookSet();
        public IWorkbook workbook;
        private SheetInfo ActiveSheetInfo { get; set; }

        public int WorksheetCount { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        public ExcelWriter(string filename)
        {
            if (File.Exists(filename))
                workbook = workbookSet.Workbooks.Open(filename);
            else if (File.Exists(filename + ".xls"))
                workbook = workbookSet.Workbooks.Open(filename + ".xls");
            else if (File.Exists(filename + ".xlsx"))
                workbook = workbookSet.Workbooks.Open(filename + ".xlsx");
            else
                CreateWorkbook(filename);
        }

        private void CreateWorkbook(string filename)
        {
            workbook = workbookSet.Workbooks.Add();
            if (filename == "" || filename == null)
                workbook.FullName = "dss_excel" + Tools.RandomString(10) + ".xlsx";
            else if (filename.EndsWith(".xls") || filename.EndsWith(".xlsx"))
                workbook.FullName = filename;
            else
            {
                workbook.FullName = Path.GetDirectoryName(filename) + "\\" +
                    Path.GetFileNameWithoutExtension(filename) + ".xlsx";
            }
        }

        private static void SetIndexColumnInExcelFile(IWorkbook book, string sheet, object record)
        {
            book.Worksheets["Sheet1"].Cells[0, 0].Value = "Index";
            int rowOffset = 1;
            if (record is TimeSeries)
            {
                var ts = (TimeSeries)record;
                for (int i = 0 + rowOffset; i < ts.Count + rowOffset; i++)
                {
                    book.Worksheets["Sheet1"].Cells[i, 0].Value = i - rowOffset + 1;

                }
            }
            else if (record is PairedData)
            {
                var pd = (PairedData)record;
                for (int i = 0 + rowOffset; i < pd.XCount + rowOffset; i++)
                {
                    book.Worksheets["Sheet1"].Cells[i, 0].Value = i - rowOffset + 1;

                }
            }
        }

        private static void SetDateColumnInExcelFile(IWorkbook book, string sheet, object record, int rowOffset, int colOffset)
        {
            if (record is TimeSeries)
            {
                book.Worksheets[sheet].Cells[rowOffset, colOffset].Value = "Date/Time";
                var ts = (TimeSeries)record;
                for (int i = 0 + rowOffset + 1; i < ts.Count + rowOffset + 1; i++)
                {
                    book.Worksheets[sheet].Cells[i, colOffset].Value = ts.Times[i - rowOffset - 1];
                }
            }
        }

        private static void SetOrdinateColumnInExcelFile(IWorkbook book, string sheet, object record, int rowOffset, int colOffset)
        {

            if (record is PairedData)
            {
                book.Worksheets[sheet].Cells[rowOffset, colOffset].Value = "Ordinates";
                var pd = (PairedData)record;
                for (int i = 0 + rowOffset + 1; i < pd.XCount + rowOffset + 1; i++)
                {
                    book.Worksheets[sheet].Cells[i, colOffset].Value = pd.Ordinates[i - rowOffset - 1];
                }

            }
        }

        public void Write(TimeSeries record, string sheet)
        {
            if (!SheetExists(sheet))
                AddSheet(sheet);
            ClearSheet(sheet);
            SetPathInExcelFile(sheet, record.Path);
            SetUnitsAndDataTypeInExcelFile(sheet, record.Units, record.DataType);
            SetDateColumnInExcelFile(workbook, sheet, record, (int)PathLayout.StandardPath, 0);
            SetTimeSeriesValueColumnInExcelFile(workbook, sheet, record, (int)PathLayout.StandardPath, 1);
            if (workbook.FullName.EndsWith(".xls"))
                workbook.SaveAs(workbook.FullName, FileFormat.Excel8);
            else if (workbook.FullName.EndsWith(".xlsx"))
                workbook.SaveAs(workbook.FullName, FileFormat.OpenXMLWorkbook);
            else
            {
                var name = Path.GetDirectoryName(workbook.FullName) + "\\" +
                    Path.GetFileNameWithoutExtension(workbook.FullName) + ".xlsx";
                workbook.SaveAs(name, FileFormat.OpenXMLWorkbook);
            }
        }

        private void SetUnitsAndDataTypeInExcelFile(string sheet, string units, string dataType)
        {
            workbook.Worksheets[sheet].Cells[6, 0].Value = "Units";
            workbook.Worksheets[sheet].Cells[6, 1].Value = units;

            workbook.Worksheets[sheet].Cells[7, 0].Value = "Data Type";
            workbook.Worksheets[sheet].Cells[7, 1].Value = dataType;
        }

        public void Write(IEnumerable<TimeSeries> records, string sheet)
        {
            if (!SheetExists(sheet))
                AddSheet(sheet);
            ClearSheet(sheet);
            SetDateColumnInExcelFile(workbook, sheet, records, (int)PathLayout.StandardPath, 0);
            SetPathUnitsAndDataTypeInExcelFile(workbook, sheet, records, 1);
            SetTimeSeriesValueColumnInExcelFile(workbook, sheet, records, (int)PathLayout.StandardPath, 1);
            if (workbook.FullName.EndsWith(".xls"))
                workbook.SaveAs(workbook.FullName, FileFormat.Excel8);
            else if (workbook.FullName.EndsWith(".xlsx"))
                workbook.SaveAs(workbook.FullName, FileFormat.OpenXMLWorkbook);
            else
            {
                var name = Path.GetDirectoryName(workbook.FullName) + "\\" +
                    Path.GetFileNameWithoutExtension(workbook.FullName) + ".xlsx";
                workbook.SaveAs(name, FileFormat.OpenXMLWorkbook);
            }
        }

        private void SetPathUnitsAndDataTypeInExcelFile(IWorkbook workbook, string sheet, IEnumerable<TimeSeries> records, int columnOffset)
        {
            workbook.Worksheets[sheet].Cells[0, 0].Value = "A";
            workbook.Worksheets[sheet].Cells[1, 0].Value = "B";
            workbook.Worksheets[sheet].Cells[2, 0].Value = "C";
            workbook.Worksheets[sheet].Cells[3, 0].Value = "D";
            workbook.Worksheets[sheet].Cells[4, 0].Value = "E";
            workbook.Worksheets[sheet].Cells[5, 0].Value = "F";

            for (int i = 0; i < records.Count(); i++)
            {
                workbook.Worksheets[sheet].Cells[0, i + columnOffset].Value = records.ElementAt(i).Path.Apart;
                workbook.Worksheets[sheet].Cells[1, i + columnOffset].Value = records.ElementAt(i).Path.Bpart;
                workbook.Worksheets[sheet].Cells[2, i + columnOffset].Value = records.ElementAt(i).Path.Cpart;
                workbook.Worksheets[sheet].Cells[3, i + columnOffset].Value = records.ElementAt(i).Path.Dpart;
                workbook.Worksheets[sheet].Cells[4, i + columnOffset].Value = records.ElementAt(i).Path.Epart;
                workbook.Worksheets[sheet].Cells[5, i + columnOffset].Value = records.ElementAt(i).Path.Fpart;
                workbook.Worksheets[sheet].Cells[6, i + columnOffset].Value = records.ElementAt(i).Units;
                workbook.Worksheets[sheet].Cells[7, i + columnOffset].Value = records.ElementAt(i).DataType;
            }
        }

        private void SetTimeSeriesValueColumnInExcelFile(IWorkbook workbook, string sheet, IEnumerable<TimeSeries> records, int rowOffset, int colOffset)
        {
            
            for (int j = 0; j < records.Count(); j++)
            {
                workbook.Worksheets[sheet].Cells[rowOffset, colOffset + j].Value = "Values " + (j + 1).ToString();
                for (int i = rowOffset + 1; i < records.ElementAt(j).Count + rowOffset + 1; i++)
                    workbook.Worksheets[sheet].Cells[i, colOffset + j].Value = records.ElementAt(j).Values[i - rowOffset - 1];
            }
        }

        private void SetTimeSeriesValueColumnInExcelFile(IWorkbook workbook, string sheet, TimeSeries ts, int rowOffset, int colOffset)
        {
            workbook.Worksheets[sheet].Cells[rowOffset, colOffset].Value = "Values";
            for (int i = rowOffset + 1; i < ts.Count + rowOffset + 1; i++)
                workbook.Worksheets[sheet].Cells[i, colOffset].Value = ts.Values[i - rowOffset - 1];
        }

        public void Write(PairedData record, string sheet)
        {
            if (!SheetExists(sheet))
                AddSheet(sheet);
            ClearSheet(sheet);
            SetPathInExcelFile(sheet, record.Path);
            SetOrdinateColumnInExcelFile(workbook, sheet, record, 6, 0);
            SetPairedDataValueColumnsInExcelFile(workbook, sheet, record, 6, 1);
            if (workbook.FullName.EndsWith(".xls"))
                workbook.SaveAs(workbook.FullName, FileFormat.Excel8);
            else if (workbook.FullName.EndsWith(".xlsx"))
                workbook.SaveAs(workbook.FullName, FileFormat.OpenXMLWorkbook);
            else
            {
                var name = Path.GetDirectoryName(workbook.FullName) + "\\" +
                    Path.GetFileNameWithoutExtension(workbook.FullName) + ".xlsx";
                workbook.SaveAs(name, FileFormat.OpenXMLWorkbook);
            }
        }

        public void ClearSheet(string sheet)
        {
            workbook.Worksheets[sheet].Cells.Clear();
        }

        private void SetPairedDataValueColumnsInExcelFile(IWorkbook workbook, string sheet, PairedData pd, int rowOffset, int colOffset)
        {
            for (int i = 0 + colOffset; i < pd.YCount + colOffset; i++)
            {
                workbook.Worksheets[sheet].Cells[rowOffset, i].Value = "Value " + (i - colOffset + 1).ToString();
            }

            for (int i = 0 + colOffset; i < pd.YCount + colOffset; i++)
            {
                for (int j = 0 + rowOffset + 1; j < pd.XCount + rowOffset + 1; j++)
                {
                    workbook.Worksheets[sheet].Cells[j, i].Value = pd.Values[i - colOffset][j - rowOffset - 1];
                }
            }
        }

        private void SetPathInExcelFile(string sheet, DssPath path)
        {
            workbook.Worksheets[sheet].Cells[0, 0].Value = "A";
            workbook.Worksheets[sheet].Cells[0, 1].Value = path.Apart;

            workbook.Worksheets[sheet].Cells[1, 0].Value = "B";
            workbook.Worksheets[sheet].Cells[1, 1].Value = path.Bpart;

            workbook.Worksheets[sheet].Cells[2, 0].Value = "C";
            workbook.Worksheets[sheet].Cells[2, 1].Value = path.Cpart;

            workbook.Worksheets[sheet].Cells[3, 0].Value = "D";
            workbook.Worksheets[sheet].Cells[3, 1].Value = path.Dpart;

            workbook.Worksheets[sheet].Cells[4, 0].Value = "E";
            workbook.Worksheets[sheet].Cells[4, 1].Value = path.Epart;

            workbook.Worksheets[sheet].Cells[5, 0].Value = "F";
            workbook.Worksheets[sheet].Cells[5, 1].Value = path.Fpart;
        }

        public void Write(TimeSeries record, int sheetIndex)
        {
            Write(record, workbook.Worksheets[sheetIndex].Name);
        }

        public void Write(PairedData record, int sheetIndex)
        {
            Write(record, workbook.Worksheets[sheetIndex].Name);
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
            return !isRegularTimeSeries(worksheet);
        }

        public bool isPairedData(string worksheet)
        {
            var offset = ActiveSheetInfo.HasIndex ? 3 : 2;
            if (ColumnCount(worksheet) < offset)
                return false;

            for (int i = ActiveSheetInfo.DataStartRowIndex; i < ActiveSheetInfo.SmallestColumnRowCount; i++)
            {
                for (int j = 0; i < ActiveSheetInfo.ColumnCount; i++)
                {
                    if (IsValue(workbook.Worksheets[worksheet].Cells[i, j]))
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

            if (!IsValue(workbook.Worksheets[worksheet].Cells[start, 0]) &&
                vals[start, 0].Number != 0 && vals[start, 0].Number != 1)
                return false;

            for (int i = start; i < end; i++)
                l.Add((int)(vals[i, 0].Number));

            return l.ToArray().SequenceEqual(Enumerable.Range(1, l.Count)) ||
                l.ToArray().SequenceEqual(Enumerable.Range(0, l.Count - 1)) ? true : false;
        }

        public bool HasDate(string worksheet)
        {
            var cells = (workbook.Worksheets[worksheet]).Cells;
            return HasIndex(worksheet) ? IsDate(cells[SmallestColumnRowCount(worksheet) - 1, 1]) : IsDate(cells[SmallestColumnRowCount(worksheet) - 1, 0]);
        }

        public int RowCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.RowCount;
        }

        public int ColumnCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.ColumnCount;
        }

        public DateTime GetDateFromCell(string s)
        {
            CorrectDateFormat(s, out DateTime dt);
            return dt;
        }

        public bool IsRegular(List<DateTime> times)
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

        public IEnumerable<TimeSeries> GetTimeSeries(IRange DateTimes, IRange Values, string Apart, string Bpart, string Cpart, string Dpart, string Epart, string Fpart)
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

        public double[] RangeToTimeSeriesValues(IRange values, int columnIndex)
        {
            var d = new List<double>();

            for (int i = 0; i < values.RowCount; i++)
            {
                d.Add(double.Parse(values[i, columnIndex].Value.ToString()));
            }

            return d.ToArray();
        }

        public DateTime[] RangeToDateTimes(IRange dateTimes)
        {
            var r = new List<DateTime>();
            for (int i = 0; i < dateTimes.RowCount; i++)
            {
                CorrectDateFormat(CellToString(dateTimes[i, 0]), out DateTime tmp);
                r.Add(tmp);
            }
            return r.ToArray();
        }

        public PairedData GetPairedData(IRange Ordinates, IRange Values, string Apart, string Bpart, string Cpart, string Dpart, string Epart, string Fpart)
        {
            var pd = new PairedData();
            pd.Ordinates = RangeToOrdinates(Ordinates);
            pd.Values = RangeToPairedDataValues(Values);
            var p = new DssPath(Apart, Bpart, Cpart, Dpart, Epart, Fpart, RecordType.PairedData, "type", "units");
            pd.Path = p;
            return pd;
        }

        public List<double[]> RangeToPairedDataValues(IRange values)
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

        public double[] RangeToOrdinates(IRange ordinates)
        {
            var d = new List<double>();

            for (int i = 0; i < ordinates.RowCount; i++)
            {
                d.Add(double.Parse(ordinates[i, 0].Value.ToString()));
            }

            return d.ToArray();
        }

        public RecordType CheckTimeSeriesType(DateTime[] times)
        {
            return IsRegular(times.ToList()) ? RecordType.RegularTimeSeries : RecordType.IrregularTimeSeries;
        }

        public bool IsDateRange(IRange range)
        {
            for (int i = 0; i < range.RowCount; i++)
            {
                if (!IsDate(range[i, 0]))
                    return false;
            }
            return true;
        }

        public bool IsDate(IRange date)
        {
            if (!IsValidCell(date))
                return false;

            CorrectDateFormat(date[0, 0].Text, out DateTime d);
            return d == new DateTime() ? false : DateTime.TryParse(d.ToString(), out _);
        }

        public bool IsValidCell(IRange cell)
        {
            if (cell[0, 0].Value == null || cell[0, 0].Text.Trim() == "")
                return false;

            return true;
        }

        public void CorrectDateFormat(string s, out DateTime d)
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

        public bool IsDifferentDateFromat(string s, out DateTime d)
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

        public bool IsOrdinateRange(IRange range)
        {
            return IsValueRange(range);
        }

        public bool IsValueRange(IRange range)
        {
            for (int i = 0; i < range.RowCount; i++)
            {
                if (!IsValue(range[i, 0]))
                    return false;
            }
            return true;
        }

        public bool IsValuesRange(IRange range)
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

        public bool IsValue(IRange value)
        {
            if (!IsValidCell(value))
                return false;

            return double.TryParse(value[0, 0].Text, out _);
        }

        public string CellToString(IRange value)
        {
            return value[0, 0].Text;
        }

        public bool IsAllColumnRowCountsEqual(IRange range)
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
            int r = RowCount(worksheet);
            int s = RowCount(worksheet) - 1;
            int c = ColumnCount(worksheet);
            IRange cells = workbook.Worksheets[worksheet].Cells;
            for (int i = 0; i < c; i++)
            {
                for (int j = r - 1; j > DataStartRowIndex(worksheet); j--)
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
    }
}
