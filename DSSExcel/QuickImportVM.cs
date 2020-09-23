using Hec.Dss;
using Hec.Dss.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace DSSExcel
{

    public class QuickImportVM : INotifyPropertyChanged
    {
        private bool hasExcelFile = false;
        private bool hasDssFile = false;
        private string excelFilePath = "";
        private string dssFilePath = "";
        private bool overwriteSheets = false;
        private List<string> paths = new List<string>();
        private List<string> sheets = new List<string>();
        public bool HasExcelFile
        {
            get { return hasExcelFile; }
            set
            {
                hasExcelFile = value;
                NotifyPropertyChanged(nameof(HasExcelFile));
            }
        }
        public bool HasDssFile
        {
            get { return hasDssFile; }
            set
            {
                hasDssFile = value;
                NotifyPropertyChanged(nameof(HasDssFile));
            }
        }
        public string ExcelFilePath
        {
            get { return excelFilePath; }
            set
            {
                excelFilePath = value;
                NotifyPropertyChanged(nameof(ExcelFilePath));
            }
        }
        public string DssFilePath
        {
            get { return dssFilePath; }
            set
            {
                dssFilePath = value;
                NotifyPropertyChanged(nameof(DssFilePath));
            }
        }
        public bool OverwriteSheets 
        { 
            get { return overwriteSheets; }
            set
            {
                overwriteSheets = value;
                NotifyPropertyChanged(nameof(OverwriteSheets));
            }
        }

        public List<string> Paths
        {
            get { return paths; }
            set
            {
                paths = value;
                NotifyPropertyChanged(nameof(Paths));
            }
        }
        public List<string> Sheets
        {
            get { return sheets; }
            set
            {
                sheets = value;
                NotifyPropertyChanged(nameof(Sheets));
            }
        }
        public List<string> SelectedPaths { get; set; }
        public List<string> SelectedSheets { get; set; }

        public QuickImportVM()
        {

        }

        public event PropertyChangedEventHandler PropertyChanged;
        public void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void QuickImport()
        {
            ExcelReader er = new ExcelReader(ExcelFilePath);
            using (DssWriter w = new DssWriter(DssFilePath))
            {

                foreach (var sheet in SelectedSheets)
                {
                    var t = er.CheckType(sheet);
                    if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                        w.Write(er.Read(sheet) as TimeSeries);
                    else if (t is RecordType.PairedData)
                        w.Write(er.Read(sheet) as PairedData);
                }
            }
            GetAllPaths();
        }

        public void QuickExport()
        {
            using (DssReader r = new DssReader(DssFilePath))
            {
                object record;
                ExcelWriter ew = new ExcelWriter(ExcelFilePath);
                for (int i = 0; i < SelectedSheets.Count; i++)
                {
                    DssPath p = new DssPath(SelectedPaths[i]);
                    var type = r.GetRecordType(p);
                    if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                    {
                        record = r.GetTimeSeries(p);
                        ew.Write(record as TimeSeries, SelectedSheets[i]);
                    }
                    else if (type is RecordType.PairedData)
                    {
                        record = r.GetPairedData(p.FullPath);
                        ew.Write(record as PairedData, SelectedSheets[i]);
                    }
                }
            }
            GetAllSheets();
        }

        public void GetAllSheets()
        {
            var s = new List<string>();
            ExcelReader er = new ExcelReader(ExcelFilePath);
            for (int i = 0; i < er.workbook.Worksheets.Count; i++)
                s.Add(er.workbook.Worksheets[i].Name);
            Sheets = s;
        }

        public void GetAllPaths()
        {
            List<string> p = new List<string>();
            using (DssReader r = new DssReader(DssFilePath))
            {
                DssPathCollection c;
                c = r.GetCatalog();
                foreach (var path in c)
                    p.Add(path.FullPath);
            }
            Paths = p;
        }
    }
}
