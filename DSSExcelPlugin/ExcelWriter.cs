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
    public class ExcelWriter : ExcelTools
    {

        public ExcelWriter(string filename)
        {
            if (!File.Exists(filename))
                CreateWorkbook(filename);
            else
                workbook = workbookSet.Workbooks.Open(filename);
        }

        private void CreateWorkbook(string filename)
        {
            var fn = Path.GetDirectoryName(filename) + "\\" + Path.GetFileNameWithoutExtension(filename) + ".xls";
            IWorkbookSet bookSet = Factory.GetWorkbookSet();
            workbook = bookSet.Workbooks.Add();
            workbook.FullName = fn;
        }

        private static void SetIndexColumnInExcelFile(IWorkbook book, string sheet, object record)
        {
            book.Worksheets["Sheet1"].Cells[0, 0].Value = "Index";
            int rowOffset = 1;
            if (record is TimeSeries)
            {
                var ts = (TimeSeries)record;
                for (int i = 0 + rowOffset; i < ts.Count + rowOffset; i++)
                {
                    book.Worksheets["Sheet1"].Cells[i, 0].Value = i - rowOffset + 1;

                }
            }
            else if (record is PairedData)
            {
                var pd = (PairedData)record;
                for (int i = 0 + rowOffset; i < pd.XCount + rowOffset; i++)
                {
                    book.Worksheets["Sheet1"].Cells[i, 0].Value = i - rowOffset + 1;

                }
            }
        }

        private static void SetDateColumnInExcelFile(IWorkbook book, string sheet, object record)
        {
            if (record is TimeSeries)
            {
                book.Worksheets[sheet].Cells[0, 0].Value = "Date/Time";
                int rowOffset = 1;
                var ts = (TimeSeries)record;
                for (int i = 0 + rowOffset; i < ts.Count + rowOffset; i++)
                {
                    book.Worksheets[sheet].Cells[i, 0].Value = ts.Times[i - rowOffset];
                }
            }
        }

        private static void SetOrdinateColumnInExcelFile(IWorkbook book, string sheet, object record)
        {

            if (record is PairedData)
            {
                book.Worksheets[sheet].Cells[0, 0].Value = "Ordinates";
                int rowOffset = 1;
                var pd = (PairedData)record;
                for (int i = 0 + rowOffset; i < pd.XCount + rowOffset; i++)
                {
                    book.Worksheets[sheet].Cells[i, 0].Value = pd.Ordinates[i - rowOffset];
                }

            }
        }

        private static void SetValueColumnInExcelFile(IWorkbook book, string sheet, object record)
        {

            if (record is TimeSeries)
            {
                var ts = (TimeSeries)record;
                book.Worksheets[sheet].Cells[0, 1].Value = "Values";
                int offset = 1;
                for (int i = 0 + offset; i < ts.Count + offset; i++)
                {
                    book.Worksheets[sheet].Cells[i, 1].Value = ts.Values[i - offset];
                }

            }
            else if (record is PairedData)
            {
                var pd = (PairedData)record;
                int columnOffset = 1;
                for (int i = 0 + columnOffset; i < pd.YCount + columnOffset; i++)
                {
                    book.Worksheets[sheet].Cells[0, i].Value = "Value " + (i).ToString();

                }
                int rowOffset = 1;

                for (int i = 0 + columnOffset; i < pd.YCount + columnOffset; i++)
                {
                    for (int j = 0 + rowOffset; j < pd.XCount + rowOffset; j++)
                    {
                        book.Worksheets[sheet].Cells[j, i].Value = pd.Values[i - columnOffset][j - rowOffset];
                    }
                }

            }
        }

        public void Write(TimeSeries record, string sheet)
        {
            SetDateColumnInExcelFile(workbook, sheet, record);
            SetValueColumnInExcelFile(workbook, sheet, record);
            if (workbook.FullName.EndsWith(".xls"))
                workbook.SaveAs(workbook.FullName, FileFormat.Excel8);
            else if (workbook.FullName.EndsWith(".xlsx"))
                workbook.SaveAs(workbook.FullName, FileFormat.OpenXMLWorkbook);
            else
            {
                var name = Path.GetDirectoryName(workbook.FullName) + "\\" +
                    Path.GetFileNameWithoutExtension(workbook.FullName) + ".xlsx";
                workbook.SaveAs(name, FileFormat.OpenXMLWorkbook);
            }
            workbook.Close();
        }

        public void Write(PairedData record, string sheet)
        {
            SetOrdinateColumnInExcelFile(workbook, sheet, record);
            SetValueColumnInExcelFile(workbook, sheet, record);
            if (workbook.FullName.EndsWith(".xls"))
                workbook.SaveAs(workbook.FullName, FileFormat.Excel8);
            else if (workbook.FullName.EndsWith(".xlsx"))
                workbook.SaveAs(workbook.FullName, FileFormat.OpenXMLWorkbook);
            else
            {
                var name = Path.GetDirectoryName(workbook.FullName) + "\\" +
                    Path.GetFileNameWithoutExtension(workbook.FullName) + ".xlsx";
                workbook.SaveAs(name, FileFormat.OpenXMLWorkbook);
            }
            workbook.Close();
        }
        
        public void Write(TimeSeries record, int sheetIndex)
        {
            Write(record, workbook.Worksheets[sheetIndex].Name);
        }

        public void Write(PairedData record, int sheetIndex)
        {
            Write(record, workbook.Worksheets[sheetIndex].Name);
        }


    }
}
