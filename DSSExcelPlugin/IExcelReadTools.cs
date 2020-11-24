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