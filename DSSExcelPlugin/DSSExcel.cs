using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Primitives;
using System.Data;
using System.Data.SqlClient;
using System.Dynamic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Hec.Dss;
using SpreadsheetGear;
using SpreadsheetGear.Advanced.Cells;

namespace DSSExcelPlugin
{
    public class DSSExcel
    {

        public SpreadsheetGear.IWorkbookSet workbookSet = SpreadsheetGear.Factory.GetWorkbookSet();
        public SpreadsheetGear.IWorkbook workbook;

        public void ChangeActiveSheet(string worksheet)
        {
            workbook.Worksheets[worksheet].Select();
        }

        public void ChangeActiveSheet(int worksheet)
        {
            workbook.Worksheets[worksheet].Select();
        }

        public DSSExcel(string filename)
        {
            workbook = workbookSet.Workbooks.Open(filename);
            ChangeActiveSheet(0);
        }

        /// <summary>
        /// Returns the row index where the headers end and the data begins.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private int DataStartIndex(string worksheet)
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

        private bool isRegularTimeSeries(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var d = new List<DateTime>();
            if (HasDate(worksheet))
            {
                if (HasIndex(worksheet))
                {
                    for (int i = DataStartIndex(worksheet); i < RowCount(worksheet); i++)
                    {
                        
                        DateTime dt = GetDateFromExcel(vals[i, 1].Number);
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

                        DateTime dt = GetDateFromExcel(vals[i, 0].Number);
                        d.Add(dt);
                    }
                    if (IsRegular(d))
                        return true;
                    return false;
                }
            }
            return false;
        }

        private bool isIrregularTimeSeries(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var d = new List<DateTime>();
            if (HasDate(worksheet))
            {
                if (HasIndex(worksheet))
                {
                    for (int i = DataStartIndex(worksheet); i < RowCount(worksheet); i++)
                    {
                        DateTime dt = GetDateFromExcel(vals[i, 1].Number);
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

                        DateTime dt = GetDateFromExcel(vals[i, 0].Number);
                        d.Add(dt);
                    }
                    if (IsRegular(d))
                        return false;
                    return true;
                }
            }
            return false;
        }

        private bool isPairedData(string worksheet)
        {
            throw new NotImplementedException();
        }

        private bool isGrid(string worksheet)
        {
            throw new NotImplementedException();
        }

        private bool isTin(string worksheet)
        {
            throw new NotImplementedException();
        }

        private bool isLocationInfo(string worksheet)
        {
            throw new NotImplementedException();
        }

        private bool isText(string worksheet)
        {
            throw new NotImplementedException();
        }

        private bool HasIndex(string worksheet)
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

        public bool HasDate(string worksheet)
        {
            var vals = (IValues)(workbook.Worksheets[worksheet]);
            if (HasIndex(worksheet))
            {
                return DateTime.TryParse(workbook.NumberToDateTime(vals[RowCount(worksheet) - 1, 1].Number).ToString(), out _);
            }
            else
            {
                return DateTime.TryParse(workbook.NumberToDateTime(vals[RowCount(worksheet) - 1, 0].Number).ToString(), out _);
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

        private int RowCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.RowCount;
        }

        private int ColumnCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.ColumnCount;
        }

        public TimeSeries DataTableToTimeSeries(DataTable dataTable)
        {
            var ts = new TimeSeries();
            


            return ts;
        }

        public PairedData DataTabletoPairedData(DataTable dataTable)
        {
            var pd = new PairedData();



            return pd;
        }

        /// <summary>
        /// Gets DateTime object from the double value of a date from an excel sheet.
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        private DateTime GetDateFromExcel(double date)
        {
            DateTime dt;
            var b = DateTime.TryParse(workbook.NumberToDateTime(date).ToString(), out dt);
            return b ? dt : new DateTime();
        }

        private bool IsRegular(List<DateTime> times)
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



    }
}
