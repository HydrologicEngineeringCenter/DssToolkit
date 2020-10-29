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
    public enum PathLayout
    {
        StandardPath = 8,
        PathWithoutDPart = 7,
        StandardPathWithoutTypeAndUnits = 6,
        StandardPathWithoutDPartTypeAndUnits = 5,
        NoPath = 0
    }

    public interface IExcelReadTools
    {
        int WorksheetCount { get; }

        /// <summary>
        /// Returns the 0-based row index where the headers end and the data begins.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        int DataStartRowIndex(string worksheet);

        int DataStartRow(string worksheet);

        RecordType CheckType(string worksheet);

        RecordType CheckType(int worksheetIndex);

        bool isRegularTimeSeries(string worksheet);

        IValues Values(string worksheet);

        bool isIrregularTimeSeries(string worksheet);

        bool isPairedData(string worksheet);

        bool isGrid(string worksheet);

        bool isTin(string worksheet);

        bool isLocationInfo(string worksheet);

        bool isText(string worksheet);

        bool HasIndex(string worksheet);

        bool HasDate(string worksheet);

        /// <summary>
        /// Returns default row count of a given worksheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        int RowCount(string worksheet);

        int ColumnCount(string worksheet);

        DateTime GetDateFromCell(string s);

        bool IsRegular(List<DateTime> times);

        void AddSheet(string sheet);

        void AddSheet(int sheetIndex);

        bool SheetExists(string sheet);

        bool SheetExists(int sheetIndex);

        IEnumerable<TimeSeries> GetTimeSeries(IRange DateTimes, IRange Values, string Apart, string Bpart, string Cpart, string Dpart, string Epart, string Fpart);

        /// <summary>
        /// Convert a specified column in a range of values to a double array.
        /// </summary>
        /// <param name="values"></param>
        /// <param name="columnIndex"></param>
        /// <returns></returns>
        double[] RangeToTimeSeriesValues(IRange values, int columnIndex);

        DateTime[] RangeToDateTimes(IRange dateTimes);

        PairedData GetPairedData(IRange Ordinates, IRange Values, string Apart, string Bpart, string Cpart, string Dpart, string Epart, string Fpart);

        List<double[]> RangeToPairedDataValues(IRange values);

        double[] RangeToOrdinates(IRange ordinates);

        RecordType CheckTimeSeriesType(DateTime[] times);

        bool IsDateRange(IRange range);

        bool IsDate(IRange date);

        bool IsValidCell(IRange cell);

        void CorrectDateFormat(string s, out DateTime d);

        bool IsDifferentDateFromat(string s, out DateTime d);

        bool IsOrdinateRange(IRange range);

        bool IsValueRange(IRange range);

        bool IsValuesRange(IRange range);

        bool IsValue(IRange value);

        string CellToString(IRange value);

        bool IsAllColumnRowCountsEqual(IRange range);

        /// <summary>
        /// Returns the smallest row count of all columns in a given worksheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        int SmallestColumnRowCount(string worksheet);
    }
}