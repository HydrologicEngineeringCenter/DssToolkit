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
        public SpreadsheetGear.IWorkbookSet workbookSet = SpreadsheetGear.Factory.GetWorkbookSet();
        public SpreadsheetGear.IWorkbook workbook;

        //public ExcelTools(string filename)
        //{
        //    workbook = workbookSet.Workbooks.Open(filename);
        //}

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
            for (int i = 0; i < r; i++)
            {
                for (int j = 0; j < c; j++)
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
            var r = RowCount(worksheet);
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

            for (int i = start; i < RowCount(worksheet); i++)
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
                return cells[RowCount(worksheet) - 1, 1].NumberFormatType == NumberFormatType.DateTime ||
                    cells[RowCount(worksheet) - 1, 1].NumberFormatType == NumberFormatType.Date;
            }
            else
            {
                return cells[RowCount(worksheet) - 1, 0].NumberFormatType == NumberFormatType.DateTime ||
                    cells[RowCount(worksheet) - 1, 0].NumberFormatType == NumberFormatType.Date;
            }
        }

        public DataTable ExcelToDataTable(string worksheet)
        {
            var r = RowCount(worksheet);
            var c = ColumnCount(worksheet);

            var vals = (IValues)(workbook.Worksheets[worksheet]);
            DataTable data = new DataTable();
            for (int i = 0; i < c; i++) { data.Columns.Add(); }
            var Row = new List<object>();

            for (int i = 0; i < r; i++)
            {
                for (int j = 0; j < c; j++)
                {
                    if (vals[i, j].Type == SpreadsheetGear.Advanced.Cells.ValueType.Number)
                    {
                        Row.Add(vals[i, j].Number);
                    }
                    else if (vals[i, j].Type == SpreadsheetGear.Advanced.Cells.ValueType.Text)
                    {
                        Row.Add(vals[i, j].Text);
                    }
                }
                data.Rows.Add(Row.ToArray());
                Row.Clear();
            }
            return data;
        }

        public DataTable ExcelToDataTable(int worksheetIndex)
        {
            return ExcelToDataTable(workbook.Worksheets[worksheetIndex].Name);
        }

        protected int RowCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.RowCount;
        }

        protected int ColumnCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.ColumnCount;
        }

        /// <summary>
        /// Gets DateTime object from the double value of a date from an excel sheet.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        protected DateTime GetDateFromCell(double value)
        {
            DateTime dt;
            var b = DateTime.TryParse(workbook.NumberToDateTime(value).ToString(), out dt);
            return b ? dt : new DateTime();
        }

        protected bool IsRegular(List<DateTime> times)
        {
            var temp = times;
            temp.Sort((a, b) => a.CompareTo(b));
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
        protected string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
