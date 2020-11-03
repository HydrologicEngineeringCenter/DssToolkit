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

        void AddSheet(string sheet);

        void AddSheet(int sheetIndex);

        bool SheetExists(string sheet);

        bool SheetExists(int sheetIndex);

        /// <summary>
        /// Returns the smallest row count of all columns in a given worksheet.
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        int SmallestColumnRowCount(string worksheet);
    }
}