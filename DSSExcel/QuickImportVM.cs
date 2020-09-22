using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSSExcel
{

    public class QuickImportVM
    {
        public bool HasExcelFile { get; set; }
        public bool HasDssFile { get; set; }
        public string ExcelFilePath { get; set; }
        public string DssFilePath { get; set; }
        public bool OverwriteSheets { get; set; }

        public QuickImportVM()
        {

        }
    }
}
