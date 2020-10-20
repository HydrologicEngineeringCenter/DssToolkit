using SpreadsheetGear;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Hec.Dss.Excel.ExcelTools;

namespace Hec.Dss.Excel
{
    public class SheetInfo
    {
        public int DataStartRow { get; private set; }
        public int DataStartRowIndex { get; private set; }
        public PathLayout PathStructure { get; private set; }
        public int PathStartRow { get; private set; }
        public int PathStartRowIndex { get; private set; }
        public int PathEndRow { get; private set; }
        public int PathEndRowIndex { get; private set; }
        public int RowCount { get; private set; }
        public int ColumnCount { get; private set; }
        public int SmallestColumnRowCount { get; private set; }
        public bool HasIndex { get; private set; }
        public bool HasDate { get; private set; }
        public bool HasPath { get; private set; }
        public SheetInfo(ExcelReader r, string sheet)
        {
            DataStartRow = r.DataStartRow(sheet);
            DataStartRowIndex = DataStartRow - 1;
            PathStructure = r.GetDSSPathLayout(sheet);
            PathStartRow = 1;
            PathStartRowIndex = PathStartRow - 1;
            PathEndRow = r.DSSPathEndRow(sheet, 0);
            PathEndRowIndex = PathEndRow - 1;
            RowCount = r.RowCount(sheet);
            ColumnCount = r.ColumnCount(sheet);
            SmallestColumnRowCount = r.SmallestColumnRowCount(sheet);
            HasIndex = r.HasIndex(sheet);
            HasDate = r.HasDate(sheet);
            HasPath = r.DSSPathExists(sheet, 0);

        }
    }
}
