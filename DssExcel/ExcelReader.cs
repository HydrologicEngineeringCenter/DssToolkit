using SpreadsheetGear;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  public class ExcelReader
  {
    string fileName;
    public IWorkbook Workbook { get; set; }
    public ExcelReader(string fileName)
    {
      this.fileName = fileName;
      Workbook = SpreadsheetGear.Factory.GetWorkbook(fileName);
    }
  }
}
