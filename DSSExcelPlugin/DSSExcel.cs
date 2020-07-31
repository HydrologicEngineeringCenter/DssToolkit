using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Primitives;
using System.Data;
using System.Data.SqlClient;
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
        public IValues vals
        {
            get
            {
                return (IValues)workbook.ActiveWorksheet;
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

        public RecordType CheckType(string worksheet)
        {
            throw new NotImplementedException();

        }

        public bool HasIndex(string workbook)
        {
            throw new NotImplementedException();
        }

        public bool HasDate(string worksheet)
        {
            var vals = (IValues)(workbook.Worksheets[worksheet]);
            throw new NotImplementedException();
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



    }
}
