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
        private SheetInfo ActiveSheetInfo { get; set; }

        /// <summary>
        /// Get sheet info for a specific sheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        private SheetInfo this[string worksheet] => GetWorksheetInfo(worksheet);

        private SheetInfo GetWorksheetInfo(string worksheet)
        {
            return new SheetInfo(this, worksheet);
        }

        public ExcelReader(string filename)
        {
            workbook = workbookSet.Workbooks.Open(filename);
        }

        public TimeSeries GetTimeSeries(string worksheet)
        {
            ActiveSheetInfo = this[worksheet];
            if (!isIrregularTimeSeries(worksheet) && !isRegularTimeSeries(worksheet))
                return new TimeSeries();

            TimeSeries ts = new TimeSeries();
            GetTimeSeriesTimes(ts, worksheet);
            ts.Values = GetTimeSeriesValues(worksheet);
            ts.Path = GetRandomTimeSeriesPath(ts, worksheet);
            ts.DataType = "type1";
            ts.Units = "unit1";

            return ts;
        }

        public IEnumerable<TimeSeries> GetMultipleTimeSeries(string worksheet)
        {
            ActiveSheetInfo = this[worksheet];
            if (!isIrregularTimeSeries(worksheet) && !isRegularTimeSeries(worksheet))
                return new List<TimeSeries>();
            var l = new List<TimeSeries>();
            var c = TimeSeriesValueColumnCount();
            for (int i = 0; i < c; i++)
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
            int dataTypeIndex = (int)ActiveSheetInfo.PathStructure - 1;
            int offset = ActiveSheetInfo.HasIndex ? valueColumn + 2 : valueColumn + 1;
            if (ActiveSheetInfo.PathStructure != PathLayout.NoPath)
                s = CellToString(workbook.Worksheets[worksheet].Cells[dataTypeIndex, offset]);
            ts.DataType = s;
        }

        private void GetTimeSeriesUnits(TimeSeries ts, string worksheet, int valueColumn)
        {
            var s = "Units";
            int unitIndex = (int)ActiveSheetInfo.PathStructure - 2;
            int offset = ActiveSheetInfo.HasIndex ? valueColumn + 2 : valueColumn + 1;
            if (ActiveSheetInfo.PathStructure != PathLayout.NoPath)
                s = CellToString(workbook.Worksheets[worksheet].Cells[unitIndex, offset]);
            ts.Units = s;
        }

        private void GetTimeSeriesPath(TimeSeries ts, string worksheet, int valueColumn)
        {
            int offset = ActiveSheetInfo.HasIndex ? valueColumn + 2 : valueColumn + 1;
            if (!ActiveSheetInfo.HasPath)
            {
                GetRandomTimeSeriesPath(ts, worksheet);
                return;
            }

            GetPath(ts, worksheet, offset, ActiveSheetInfo.PathStructure);
        }

        private void GetPath(TimeSeries ts, string worksheet, int column, PathLayout lastPathEntry)
        {
            ts.Path = new DssPath();

            if (lastPathEntry == PathLayout.StandardPath || lastPathEntry == PathLayout.StandardPathWithoutTypeAndUnits)
            {
                ts.Path.Apart = CellToString(workbook.Worksheets[worksheet].Cells[0, column]);
                ts.Path.Bpart = CellToString(workbook.Worksheets[worksheet].Cells[1, column]);
                ts.Path.Cpart = CellToString(workbook.Worksheets[worksheet].Cells[2, column]);
                ts.Path.Fpart = CellToString(workbook.Worksheets[worksheet].Cells[5, column]);
            }
            else if (lastPathEntry == PathLayout.PathWithoutDPart || lastPathEntry == PathLayout.StandardPathWithoutDPartTypeAndUnits)
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
            int offset = ActiveSheetInfo.HasIndex ? valueColumn + 2 : valueColumn + 1;
            for (int i = ActiveSheetInfo.DataStartRowIndex; i < ActiveSheetInfo.SmallestColumnRowCount; i++)
                v.Add(vals[i, offset].Number);
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
            var vals = Values(worksheet);
            var v = new List<double>();
            int offset = ActiveSheetInfo.HasIndex ? 2 : 1;
            for (int i = ActiveSheetInfo.DataStartRowIndex; i < ActiveSheetInfo.SmallestColumnRowCount; i++)
                v.Add(vals[i, offset].Number);
            return v.ToArray();
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
            ActiveSheetInfo = this[worksheet];
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
            var vals = Values(worksheet);
            var t = new List<double>();
            var v = new List<double[]>();
            var offset = ActiveSheetInfo.HasIndex ? 2 : 1;
            for (int i = offset; i < ActiveSheetInfo.ColumnCount; i++)
            {
                for (int j = ActiveSheetInfo.DataStartRowIndex; j < ActiveSheetInfo.SmallestColumnRowCount; j++)
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

        public int DSSPathEndRowIndex(string worksheet, int column)
        {
            int headerRow = 1;
            int dataStartRow = 1;
            return DataStartRowIndex(worksheet) - headerRow - dataStartRow; // remove the header and data start rows from data start index to get path end index
        }

        public int DSSPathEndRow(string worksheet, int column)
        {
            return DSSPathEndRowIndex(worksheet, column) + 1; // remove the header and data start rows from data start row to get path end row
        }

        public bool DSSPathExists(string worksheet, int column)
        {
            if (ActiveSheetInfo.PathEndRow < (int)PathLayout.StandardPathWithoutDPartTypeAndUnits || 
                ActiveSheetInfo.PathEndRow > (int)PathLayout.StandardPath)
                return false;
            int blankEntries = 0;
            for (int i = 0; i < ActiveSheetInfo.PathEndRow; i++) // check if all entries are blank
            {
                if (!IsValidCell(workbook.Worksheets[worksheet].Cells[i, column]))
                    blankEntries++;
            }

            return blankEntries < (int)PathLayout.StandardPathWithoutDPartTypeAndUnits; // return true if blank entries is less than the amount of entries for a minimal path
        }

        public PathLayout GetDSSPathLayout(string worksheet)
        {
            return (PathLayout)DSSPathEndRow(worksheet, 0);
        }
    }
}
