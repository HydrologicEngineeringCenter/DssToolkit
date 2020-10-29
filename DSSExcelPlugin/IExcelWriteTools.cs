using SpreadsheetGear;
using System.Collections.Generic;

namespace Hec.Dss.Excel
{
    public interface IExcelWriteTools
    {
        int WorksheetCount { get; }
        void CreateWorkbook(string filename);
        void SetIndexColumnInExcelFile(string sheet, object record);
        void SetDateColumnInExcelFile(string sheet, object record, int rowOffset, int colOffset);
        void SetOrdinateColumnInExcelFile(string sheet, object record, int rowOffset, int colOffset);
        void Write(TimeSeries record, string sheet);
        void SetUnitsAndDataTypeInExcelFile(string sheet, string units, string dataType);
        void Write(IEnumerable<TimeSeries> records, string sheet);
        void SetPathUnitsAndDataTypeInExcelFile(string sheet, IEnumerable<TimeSeries> records, int columnOffset);
        void SetTimeSeriesValueColumnInExcelFile(string sheet, IEnumerable<TimeSeries> records, int rowOffset, int colOffset);
        void SetTimeSeriesValueColumnInExcelFile(string sheet, TimeSeries ts, int rowOffset, int colOffset);
        void Write(PairedData record, string sheet);
        void ClearSheet(string sheet);
        void SetPairedDataValueColumnsInExcelFile(string sheet, PairedData pd, int rowOffset, int colOffset);
        void SetPathInExcelFile(string sheet, DssPath path);
        void Write(TimeSeries record, int sheetIndex);
        void Write(PairedData record, int sheetIndex);
        void AddSheet(string sheet);
        void AddSheet(int sheetIndex);
        bool SheetExists(string sheet);
        bool SheetExists(int sheetIndex);
    }
}