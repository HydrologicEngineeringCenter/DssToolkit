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
    public class ExcelWriter
    {
        public IWorkbookSet workbookSet = Factory.GetWorkbookSet();
        public IWorkbook workbook;
        public int WorksheetCount
        {
            get
            {
                return workbook.Worksheets.Count;
            }
        }

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

        public void CreateWorkbook(string filename)
        {
            workbook = workbookSet.Workbooks.Add();
            if (filename == "" || filename == null)
                workbook.FullName = "dss_excel" + ExcelTools.RandomString(10) + ".xlsx";
            else if (filename.EndsWith(".xls") || filename.EndsWith(".xlsx"))
                workbook.FullName = filename;
            else
            {
                workbook.FullName = Path.GetDirectoryName(filename) + "\\" +
                    Path.GetFileNameWithoutExtension(filename) + ".xlsx";
            }
        }

        private void SetIndexColumnInExcelFile(string sheet, object record)
        {
            workbook.Worksheets["Sheet1"].Cells[0, 0].Value = "Index";
            int rowOffset = 1;
            int count = 0;

            if (record is TimeSeries)
                count = ((TimeSeries)record).Count;
            else if (record is PairedData)
                count = ((PairedData)record).XCount;
            
            for (int i = 0 + rowOffset; i < count + rowOffset; i++)
                workbook.Worksheets["Sheet1"].Cells[i, 0].Value = i - rowOffset + 1;
        }

        private void SetDateColumnInExcelFile(string sheet, TimeSeries record, int rowOffset, int colOffset)
        {
            workbook.Worksheets[sheet].Cells[rowOffset, colOffset].Value = "Date/Time";
            for (int i = 0 + rowOffset + 1; i < record.Times.Length + rowOffset + 1; i++)
            {
                workbook.Worksheets[sheet].Cells[i, colOffset].Value = record.Times[i - rowOffset - 1];
            }
        }

        private void SetOrdinateColumnInExcelFile(string sheet, PairedData record, int rowOffset, int colOffset)
        {
            workbook.Worksheets[sheet].Cells[rowOffset, colOffset].Value = "Ordinates";
            for (int i = 0 + rowOffset + 1; i < record.Ordinates.Length + rowOffset + 1; i++)
            {
                workbook.Worksheets[sheet].Cells[i, colOffset].Value = record.Ordinates[i - rowOffset - 1];
            }
        }

        public void Write(TimeSeries record, string sheet)
        {
            if (!SheetExists(sheet))
                AddSheet(sheet);
            ClearSheet(sheet);
            SetPathInExcelFile(sheet, record.Path);
            SetUnitAndDataTypeInExcelFile(sheet, record);
            SetDateColumnInExcelFile(sheet, record, (int)PathLayout.TS_StandardPath, 0);
            SetTimeSeriesValueColumnInExcelFile(sheet, record, (int)PathLayout.TS_StandardPath, 1);
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

        private void SetUnitAndDataTypeInExcelFile(string sheet, TimeSeries record)
        {
            workbook.Worksheets[sheet].Cells[(int)PathLayout.TS_PathWithoutTypeAndUnits, 0].Value = "Units";
            workbook.Worksheets[sheet].Cells[(int)PathLayout.TS_PathWithoutTypeAndUnits, 1].Value = record.Units;

            workbook.Worksheets[sheet].Cells[(int)PathLayout.TS_PathWithoutTypeAndUnits + 1, 0].Value = "Data Type";
            workbook.Worksheets[sheet].Cells[(int)PathLayout.TS_PathWithoutTypeAndUnits + 1, 1].Value = record.DataType;
        }

        public void Write(IEnumerable<TimeSeries> records, string sheet)
        {
            if (!SheetExists(sheet))
                AddSheet(sheet);
            ClearSheet(sheet);
            SetPathsInExcelFile(sheet, records, 1);
            SetUnitsAndDataTypeInExcelFile(sheet, records, 1);
            SetDateColumnInExcelFile(sheet, records.ElementAt(0), (int)PathLayout.TS_StandardPath, 0);
            SetTimeSeriesValueColumnInExcelFile(sheet, records, (int)PathLayout.TS_StandardPath, 1);
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

        private void SetUnitsAndDataTypeInExcelFile(string sheet, IEnumerable<TimeSeries> records, int columnOffset)
        {
            workbook.Worksheets[sheet].Cells[6, 0].Value = "Units";
            workbook.Worksheets[sheet].Cells[7, 0].Value = "Data Type";

            for (int i = 0; i < records.Count(); i++)
            {
                workbook.Worksheets[sheet].Cells[6, i + columnOffset].Value = records.ElementAt(i).Units;
                workbook.Worksheets[sheet].Cells[7, i + columnOffset].Value = records.ElementAt(i).DataType;
            }
        }

        private void SetPathsInExcelFile(string sheet, IEnumerable<TimeSeries> records, int columnOffset)
        {
            workbook.Worksheets[sheet].Cells[0, 0].Value = "A";
            workbook.Worksheets[sheet].Cells[1, 0].Value = "B";
            workbook.Worksheets[sheet].Cells[2, 0].Value = "C";
            workbook.Worksheets[sheet].Cells[3, 0].Value = "D";
            workbook.Worksheets[sheet].Cells[4, 0].Value = "E";
            workbook.Worksheets[sheet].Cells[5, 0].Value = "F";

            for (int i = 0; i < records.Count(); i++)
            {
                workbook.Worksheets[sheet].Cells[0, i + columnOffset].Value = records.ElementAt(i).Path.Apart;
                workbook.Worksheets[sheet].Cells[1, i + columnOffset].Value = records.ElementAt(i).Path.Bpart;
                workbook.Worksheets[sheet].Cells[2, i + columnOffset].Value = records.ElementAt(i).Path.Cpart;
                workbook.Worksheets[sheet].Cells[3, i + columnOffset].Value = records.ElementAt(i).Path.Dpart;
                workbook.Worksheets[sheet].Cells[4, i + columnOffset].Value = records.ElementAt(i).Path.Epart;
                workbook.Worksheets[sheet].Cells[5, i + columnOffset].Value = records.ElementAt(i).Path.Fpart;
            }
        }

        private void SetTimeSeriesValueColumnInExcelFile(string sheet, IEnumerable<TimeSeries> records, int rowOffset, int colOffset)
        {
            
            for (int j = 0; j < records.Count(); j++)
            {
                workbook.Worksheets[sheet].Cells[rowOffset, colOffset + j].Value = "Values " + (j + 1).ToString();
                for (int i = rowOffset + 1; i < records.ElementAt(j).Count + rowOffset + 1; i++)
                    workbook.Worksheets[sheet].Cells[i, colOffset + j].Value = records.ElementAt(j).Values[i - rowOffset - 1];
            }
        }

        private void SetTimeSeriesValueColumnInExcelFile(string sheet, TimeSeries ts, int rowOffset, int colOffset)
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
            SetPathInExcelFile(sheet, record.Path);
            SetUnitsAndDataTypeInExcelFile(sheet, record);
            SetOrdinateColumnInExcelFile(sheet, record, (int)PathLayout.PD_StandardPath, 0);
            SetPairedDataValueColumnsInExcelFile(sheet, record, (int)PathLayout.PD_StandardPath, 1);
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

        private void SetUnitsAndDataTypeInExcelFile(string sheet, PairedData record)
        {
            workbook.Worksheets[sheet].Cells[(int)PathLayout.PD_PathWithoutTypesAndUnits, 0].Value = "Unit 1";
            workbook.Worksheets[sheet].Cells[(int)PathLayout.PD_PathWithoutTypesAndUnits + 1, 0].Value = "Unit 2";
            workbook.Worksheets[sheet].Cells[(int)PathLayout.PD_PathWithoutTypesAndUnits + 2, 0].Value = "Data Type 1";
            workbook.Worksheets[sheet].Cells[(int)PathLayout.PD_PathWithoutTypesAndUnits + 3, 0].Value = "Data Type 2";

            workbook.Worksheets[sheet].Cells[(int)PathLayout.PD_PathWithoutTypesAndUnits, 1].Value = record.UnitsIndependent;
            workbook.Worksheets[sheet].Cells[(int)PathLayout.PD_PathWithoutTypesAndUnits + 1, 1].Value = record.UnitsDependent;
            workbook.Worksheets[sheet].Cells[(int)PathLayout.PD_PathWithoutTypesAndUnits + 2, 1].Value = record.TypeIndependent;
            workbook.Worksheets[sheet].Cells[(int)PathLayout.PD_PathWithoutTypesAndUnits + 3, 1].Value = record.TypeDependent;
        }

        public void ClearSheet(string sheet)
        {
            workbook.Worksheets[sheet].Cells.Clear();
        }

        private void SetPairedDataValueColumnsInExcelFile(string sheet, PairedData pd, int rowOffset, int colOffset)
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

        private void SetPathInExcelFile(string sheet, DssPath path)
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

        public void AddSheet(string sheet)
        {
            var s = workbook.Worksheets.Add();
            s.Name = sheet;
        }

        public void AddSheet(int sheetIndex)
        {
            var s = workbook.Worksheets.Add();
            if (!SheetExists(sheetIndex))
                AddSheet(sheetIndex);
        }

        public bool SheetExists(string sheet)
        {
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                if (workbook.Worksheets[i].Name.ToLower() == sheet.ToLower())
                    return true;
            }
            return false;
        }

        public bool SheetExists(int sheetIndex)
        {
            return sheetIndex >= 0 && sheetIndex < WorksheetCount;
        }
    }
}
