using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Primitives;
using System.Data;
using System.Data.SqlClient;
using System.Dynamic;
using System.IO;
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

        public string FileName 
        {  
            get
            {
                return workbook.Name;
            }
        }

        public string FullName
        {
            get
            {
                return workbook.FullName;
            }
        }

        

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

        public RecordType CheckType(int worksheetIndex)
        {
            return CheckType(workbook.Worksheets[worksheetIndex].Name);
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

        private bool isPairedData(string worksheet)
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

        private bool HasDate(string worksheet)
        {
            var cells = (workbook.Worksheets[worksheet]).Cells;
            if (HasIndex(worksheet))
            {
                return cells[RowCount(worksheet) - 1, 1].NumberFormatType == NumberFormatType.DateTime;
            }
            else
            {
                return cells[RowCount(worksheet) - 1, 0].NumberFormatType == NumberFormatType.DateTime;
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

        private int RowCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.RowCount;
        }

        private int ColumnCount(string worksheet)
        {
            return workbook.Worksheets[worksheet].Cells.CurrentRegion.ColumnCount;
        }

        /// <summary>
        /// Gets DateTime object from the double value of a date from an excel sheet.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private DateTime GetDateFromCell(double value)
        {
            DateTime dt;
            var b = DateTime.TryParse(workbook.NumberToDateTime(value).ToString(), out dt);
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

        public void Import(string destination, string worksheet)
        {
            var t = CheckType(worksheet);
            if (t == RecordType.RegularTimeSeries)
                ImportRegularTimeSeries(destination, worksheet);
            else if (t == RecordType.IrregularTimeSeries)
                ImportIrregularTimeSeries(destination, worksheet);
            else if (t == RecordType.PairedData)
                ImportPairedData(destination, worksheet);
            else
                return;
        }

        public void Import(string destination, int worksheetIndex)
        {
            Import(destination, workbook.Worksheets[worksheetIndex].Name);
        }

        private void ImportRegularTimeSeries(string destination, string worksheet)
        {
            string fileName = destination;
            File.Delete(fileName);
            TimeSeries ts = GetTimeSeries(worksheet);
            ts.Path = new DssPath("excel", "import", "plugin", "", "", "regularTimeSeries" + RandomString(3));
            ts.Path.Epart = TimeWindow.GetInterval(ts);
            using (DssWriter w = new DssWriter(fileName))
            {
                w.Write(ts);
            }
        }

        private void ImportIrregularTimeSeries(string destination, string worksheet)
        {
            string fileName = destination;
            File.Delete(fileName);
            TimeSeries ts = GetTimeSeries(worksheet);
            ts.Path = new DssPath("excel", "import", "plugin", "", "IR-Year", "irregularTimeSeries" + RandomString(3));
            using (DssWriter w = new DssWriter(fileName))
            {
                w.Write(ts);

            }

        }

        private void ImportPairedData(string destination, string worksheet)
        {
            string fileName = destination;
            File.Delete(fileName);
            PairedData pd = GetPairedData(worksheet);
            using (DssWriter w = new DssWriter(fileName))
            {
                w.Write(pd);

            }

        }

        public TimeSeries GetTimeSeries(string worksheet)
        {
            TimeSeries ts = new TimeSeries();

            DateTime[] times = GetTimeSeriesTimes(worksheet);
            double[] vals = GetTimeSeriesValues(worksheet);
            ts.DataType = "";
            ts.Units = "";

            ts.Times = times;
            ts.Values = vals;

            return ts;
        }

        public TimeSeries GetTimeSeries(int worksheetIndex)
        {
            return GetTimeSeries(workbook.Worksheets[worksheetIndex].Name);
        }

        private double[] GetTimeSeriesValues(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var r = RowCount(worksheet);
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
            var r = RowCount(worksheet);
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
            double[] ordinates = GetPairedDataOrdinates(worksheet);
            List<double[]> vals = GetPairedDataValues(worksheet);
            PairedData pd = new PairedData(ordinates, vals, new List<string>(), "", "", "", "", "/excel/import/plugin//e/pairedData" + RandomString(3));

            return pd;
        }

        public PairedData GetPairedData(int worksheetIndex)
        {
            return GetPairedData(workbook.Worksheets[worksheetIndex].Name);
        }

        private double[] GetPairedDataOrdinates(string worksheet)
        {
            var vals = (IValues)workbook.Worksheets[worksheet];
            var r = RowCount(worksheet);
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
            var r = RowCount(worksheet);
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

        public static void Export(string fileName, object record)
        {
            IWorkbookSet bookSet = Factory.GetWorkbookSet();
            IWorkbook book = bookSet.Workbooks.Add();

            if (fileName.EndsWith(".xls"))
            {
                book.SaveAs(fileName, FileFormat.Excel8);
            }
            else 
            {
                book.SaveAs(fileName + ".xls", FileFormat.Excel8);
            } 

            SetIndexColumnInExcelFile(book, record);
            SetDateColumnInExcelFile(book, record);
            SetOrdinateColumnInExcelFile(book, record);
            SetValueColumnInExcelFile(book, record);
            book.Save();
            book.Close();
        }

        private static void SetIndexColumnInExcelFile(IWorkbook book, object record)
        {
            book.Worksheets["Sheet1"].Cells[0, 0].Value = "Index";
            int offset = 1;
            if (record is TimeSeries)
            {
                var ts = (TimeSeries)record;
                for (int i = 0 + offset; i < ts.Count + offset; i++)
                {
                    book.Worksheets["Sheet1"].Cells[i, 0].Value = i - offset + 1;

                }
            }
            else if (record is PairedData)
            {
                var pd = (PairedData)record;
                for (int i = 0 + offset; i < pd.XCount + offset; i++)
                {
                    book.Worksheets["Sheet1"].Cells[i, 0].Value = i - offset + 1;

                }
            }
        }

        private static void SetDateColumnInExcelFile(IWorkbook book, object record)
        {
            book.Worksheets["Sheet1"].Cells[0, 1].Value = "Date/Time";
            int offset = 1;
            if (record is TimeSeries)
            {
                var ts = (TimeSeries)record;
                for (int i = 0 + offset; i < ts.Count + offset; i++)
                {
                    book.Worksheets["Sheet1"].Cells[i, 1].Value = ts.Times[i - offset];
                }
            }
        }

        private static void SetOrdinateColumnInExcelFile(IWorkbook book, object record)
        {
            book.Worksheets["Sheet1"].Cells[0, 1].Value = "Ordinates";
            int offset = 1;
            if (record is PairedData)
            {
                var pd = (PairedData)record;
                for (int i = 0 + offset; i < pd.XCount + offset; i++)
                {
                    book.Worksheets["Sheet1"].Cells[i, 1].Value = pd.Ordinates[i - offset];
                }

            }
        }

        private static void SetValueColumnInExcelFile(IWorkbook book, object record)
        {
            
            if (record is TimeSeries)
            {
                var ts = (TimeSeries)record;
                book.Worksheets["Sheet1"].Cells[0, 2].Value = "Values";
                int offset = 1;
                for (int i = 0 + offset; i < ts.Count + offset; i++)
                {
                    book.Worksheets["Sheet1"].Cells[i, 2].Value = ts.Values[i - offset];
                }

            }
            else if (record is PairedData)
            {
                var pd = (PairedData)record;
                for (int i = 2; i < pd.YCount; i++)
                {
                    book.Worksheets["Sheet1"].Cells[0, i].Value = "Value " + (i - 1).ToString();

                }
                int offset = 1;

                for (int i = 2; i < pd.YCount; i++)
                {
                    for (int j = 0 + offset; j < pd.XCount + offset; j++)
                    {
                        book.Worksheets["Sheet1"].Cells[j, i].Value = pd.Values[i][j - offset];
                    }
                }

            }
        }

        private static Random random = new Random();
        private string RandomString(int length)
        {
            const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            return new string(Enumerable.Repeat(chars, length)
              .Select(s => s[random.Next(s.Length)]).ToArray());
        }
    }
}
