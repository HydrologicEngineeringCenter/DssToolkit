using SpreadsheetGear;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Hec.Dss.Excel;

namespace Hec.Dss.Excel
{
    public class SheetInfo
    {
        public string Name { get; set; }
        public int DataStartRow { get; private set; }
        public int DataStartRowIndex { get; private set; }
        public PathLayout PathLayout { get; private set; }
        public int PathStartRow { get; private set; }
        public int PathStartRowIndex { get; private set; }
        public int PathEndRow { get; private set; }
        public int PathEndRowIndex { get; private set; }
        public int RowCount { get; private set; }
        public int ColumnCount { get; private set; }
        public int ValueStartColumnIndex { get; private set; }
        public int SmallestColumnRowCount { get; private set; }
        public bool HasIndex { get; private set; }
        public bool HasDate { get; private set; }
        public bool HasPath { get; private set; }
        public bool HasHeaders { get; private set; }
        public RecordType RecordType { get; private set; }
        public SheetInfo(ExcelReader r, string sheet)
        {
            Name = sheet;
            DataStartRow = r.DataStartRow(sheet);
            DataStartRowIndex = DataStartRow - 1;
            PathLayout = r.GetDSSPathLayout(sheet);
            PathStartRow = PathLayout == PathLayout.NoPath ? -1 : 1;
            PathStartRowIndex = PathStartRow == -1 ? -1 : 0;
            PathEndRow = PathStartRowIndex == -1 ? -1 : r.DSSPathEndRow(sheet);
            PathEndRowIndex = PathEndRow == -1 ? -1 : PathEndRow - 1;
            RowCount = r.RowCount(sheet);
            ColumnCount = r.ColumnCount(sheet);
            SmallestColumnRowCount = r.SmallestColumnRowCount(sheet);
            HasIndex = r.HasIndex(sheet);
            HasDate = r.HasDate(sheet);
            HasPath = r.DSSPathExists(sheet, 0);
            ValueStartColumnIndex = HasIndex ? 2 : 1;
            HasHeaders = DataStartRowIndex != 0 && PathEndRow != DataStartRow - 1;
        }
    }
}
