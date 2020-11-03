using SpreadsheetGear;
using System.Collections.Generic;

namespace Hec.Dss.Excel
{
    public interface IExcelWriteTools
    {
        int WorksheetCount { get; }
        void CreateWorkbook(string filename);
        void Write(TimeSeries record, string sheet);
        void Write(IEnumerable<TimeSeries> records, string sheet);
        void Write(PairedData record, string sheet);
        void ClearSheet(string sheet);
        void Write(TimeSeries record, int sheetIndex);
        void Write(PairedData record, int sheetIndex);
        void AddSheet(string sheet);
        void AddSheet(int sheetIndex);
        bool SheetExists(string sheet);
        bool SheetExists(int sheetIndex);
    }
}