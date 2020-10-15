using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SpreadsheetGear;
using SpreadsheetGear.Advanced.Cells;
using SpreadsheetGear.Shapes;

namespace Hec.Dss.Excel
{
    public class ExcelWriter : ExcelTools
    {

        public ExcelWriter(string filename)
        {
            if (File.Exists(filename))
                workbook = workbookSet.Workbooks.Open(filename);
            else if (File.Exists(filename + ".xls"))
                workbook = workbookSet.Workbooks.Open(filename + ".xls");
            else if (File.Exists(filename + ".xlsx"))
                workbook = workbookSet.Workbooks.Open(filename + ".xlsx");
            else
                CreateWorkbook(filename);
        }

        private void CreateWorkbook(string filename)
        {
            workbook = workbookSet.Workbooks.Add();
            if (filename == "" || filename == null)
                workbook.FullName = "dss_excel" + RandomString(10) + ".xlsx";
            else if (filename.EndsWith(".xls") || filename.EndsWith(".xlsx"))
                workbook.FullName = filename;
            else
            {
                workbook.FullName = Path.GetDirectoryName(filename) + "\\" +
                    Path.GetFileNameWithoutExtension(filename) + ".xlsx";
            }
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

        private static void SetDateColumnInExcelFile(IWorkbook book, string sheet, object record, int rowOffset, int colOffset)
        {
            if (record is TimeSeries)
            {
                book.Worksheets[sheet].Cells[rowOffset, colOffset].Value = "Date/Time";
                var ts = (TimeSeries)record;
                for (int i = 0 + rowOffset + 1; i < ts.Count + rowOffset + 1; i++)
                {
                    book.Worksheets[sheet].Cells[i, colOffset].Value = ts.Times[i - rowOffset - 1];
                }
            }
        }

        private static void SetOrdinateColumnInExcelFile(IWorkbook book, string sheet, object record, int rowOffset, int colOffset)
        {

            if (record is PairedData)
            {
                book.Worksheets[sheet].Cells[rowOffset, colOffset].Value = "Ordinates";
                var pd = (PairedData)record;
                for (int i = 0 + rowOffset + 1; i < pd.XCount + rowOffset + 1; i++)
                {
                    book.Worksheets[sheet].Cells[i, colOffset].Value = pd.Ordinates[i - rowOffset - 1];
                }

            }
        }

        public void Write(TimeSeries record, string sheet)
        {
            if (!SheetExists(sheet))
                AddSheet(sheet);
            ClearSheet(sheet);
            SetPathInExcelFile(workbook, sheet, record.Path);
            SetDateColumnInExcelFile(workbook, sheet, record, 6, 0);
            SetTimeSeriesValueColumnInExcelFile(workbook, sheet, record, 6, 1);
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
        }

        public void Write(IEnumerable<TimeSeries> records, string sheet)
        {
            if (!SheetExists(sheet))
                AddSheet(sheet);
            ClearSheet(sheet);
            SetDateColumnInExcelFile(workbook, sheet, records, 6, 0);
            for (int i = 0; i < records.Count(); i++)
            {
                SetPathInExcelFile(workbook, sheet, records, i);
                SetTimeSeriesValueColumnInExcelFile(workbook, sheet, records, 6, 1);
            }
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
        }

        private void SetPathInExcelFile(IWorkbook workbook, string sheet, IEnumerable<TimeSeries> records, int columnOffset)
        {
            workbook.Worksheets[sheet].Cells[0, 0].Value = "A";
            workbook.Worksheets[sheet].Cells[1, 0].Value = "B";
            workbook.Worksheets[sheet].Cells[2, 0].Value = "C";
            workbook.Worksheets[sheet].Cells[3, 0].Value = "D";
            workbook.Worksheets[sheet].Cells[4, 0].Value = "E";
            workbook.Worksheets[sheet].Cells[5, 0].Value = "F";

            for (int i = 0; i < records.Count(); i++)
            {
                workbook.Worksheets[sheet].Cells[0, i + 1].Value = records.ElementAt(i).Path.Apart;

                workbook.Worksheets[sheet].Cells[1, i + 1].Value = records.ElementAt(i).Path.Bpart;

                workbook.Worksheets[sheet].Cells[2, i + 1].Value = records.ElementAt(i).Path.Cpart;

                workbook.Worksheets[sheet].Cells[3, i + 1].Value = records.ElementAt(i).Path.Dpart;

                workbook.Worksheets[sheet].Cells[4, i + 1].Value = records.ElementAt(i).Path.Epart;

                workbook.Worksheets[sheet].Cells[5, i + 1].Value = records.ElementAt(i).Path.Fpart;
            }
        }

        private void SetTimeSeriesValueColumnInExcelFile(IWorkbook workbook, string sheet, IEnumerable<TimeSeries> records, int rowOffset, int colOffset)
        {
            
            for (int j = 0; j < records.Count(); j++)
            {
                workbook.Worksheets[sheet].Cells[rowOffset, colOffset + j].Value = "Values " + (j + 1).ToString();
                for (int i = rowOffset + 1; i < records.ElementAt(j).Count + rowOffset + 1; i++)
                    workbook.Worksheets[sheet].Cells[i, colOffset + j].Value = records.ElementAt(j).Values[i - rowOffset - 1];
            }
        }

        private void SetTimeSeriesValueColumnInExcelFile(IWorkbook workbook, string sheet, TimeSeries ts, int rowOffset, int colOffset)
        {
            workbook.Worksheets[sheet].Cells[rowOffset, colOffset].Value = "Values";
            for (int i = rowOffset + 1; i < ts.Count + rowOffset + 1; i++)
                workbook.Worksheets[sheet].Cells[i, colOffset].Value = ts.Values[i - rowOffset - 1];
        }

        public void Write(PairedData record, string sheet)
        {
            if (!SheetExists(sheet))
                AddSheet(sheet);
            ClearSheet(sheet);
            SetPathInExcelFile(workbook, sheet, record.Path);
            SetOrdinateColumnInExcelFile(workbook, sheet, record, 6, 0);
            SetPairedDataValueColumnsInExcelFile(workbook, sheet, record, 6, 1);
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
        }

        public void ClearSheet(string sheet)
        {
            workbook.Worksheets[sheet].Cells.Clear();
        }

        private void SetPairedDataValueColumnsInExcelFile(IWorkbook workbook, string sheet, PairedData pd, int rowOffset, int colOffset)
        {
            for (int i = 0 + colOffset; i < pd.YCount + colOffset; i++)
            {
                workbook.Worksheets[sheet].Cells[rowOffset, i].Value = "Value " + (i - colOffset + 1).ToString();
            }

            for (int i = 0 + colOffset; i < pd.YCount + colOffset; i++)
            {
                for (int j = 0 + rowOffset + 1; j < pd.XCount + rowOffset + 1; j++)
                {
                    workbook.Worksheets[sheet].Cells[j, i].Value = pd.Values[i - colOffset][j - rowOffset - 1];
                }
            }
        }

        private void SetPathInExcelFile(IWorkbook workbook, string sheet, DssPath path)
        {
            
                workbook.Worksheets[sheet].Cells[0, 0].Value = "A";
                workbook.Worksheets[sheet].Cells[0, 1].Value = path.Apart;

                workbook.Worksheets[sheet].Cells[1, 0].Value = "B";
                workbook.Worksheets[sheet].Cells[1, 1].Value = path.Bpart;

                workbook.Worksheets[sheet].Cells[2, 0].Value = "C";
                workbook.Worksheets[sheet].Cells[2, 1].Value = path.Cpart;

                workbook.Worksheets[sheet].Cells[3, 0].Value = "D";
                workbook.Worksheets[sheet].Cells[3, 1].Value = path.Dpart;

                workbook.Worksheets[sheet].Cells[4, 0].Value = "E";
                workbook.Worksheets[sheet].Cells[4, 1].Value = path.Epart;

                workbook.Worksheets[sheet].Cells[5, 0].Value = "F";
                workbook.Worksheets[sheet].Cells[5, 1].Value = path.Fpart;

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
