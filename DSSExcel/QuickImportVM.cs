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
        private string dataFilePath = "";
        private string dssFilePath = "";
        private List<string> paths = new List<string>();
        private List<string> sheets = new List<string>();
        public bool HasDataFile { get; private set; }
        public bool HasDssFile { get; private set; }
        public string DataFilePath
        {
            get { return dataFilePath; }
            set
            {
                dataFilePath = value;
                HasDataFile = true;
                NotifyPropertyChanged(nameof(DataFilePath));
                NotifyPropertyChanged(nameof(HasDataFile));
            }
        }
        public string DssFilePath
        {
            get { return dssFilePath; }
            set
            {
                dssFilePath = value;
                HasDssFile = true;
                NotifyPropertyChanged(nameof(DssFilePath));
                NotifyPropertyChanged(nameof(HasDssFile));
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
            using (DssWriter w = new DssWriter(DssFilePath))
                ImportRecords(new ExcelReader(DataFilePath), w);
            
            GetAllPaths();
        }

        private void ImportRecords(ExcelReader er, DssWriter w)
        {
            foreach (var sheet in SelectedSheets)
            {
                var t = er.CheckType(sheet);
                if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                {
                    List<TimeSeries> l = er.ReadAll(sheet) as List<TimeSeries>;
                    foreach (var record in l)
                        w.Write(record);
                }
                else if (t is RecordType.PairedData)
                    w.Write(er.Read(sheet) as PairedData);
            }
        }

        public void QuickExport()
        {
            using (DssReader r = new DssReader(DssFilePath))
                ExportRecords(GetRecords(r), new ExcelWriter(DataFilePath));

            GetAllSheets();
        }

        private List<object> GetRecords(DssReader r)
        {
            List<object> records = new List<object>();
            for (int i = 0; i < SelectedPaths.Count; i++)
            {
                if (r.GetRecordType(new DssPath(SelectedPaths[i])) == RecordType.RegularTimeSeries ||
                    r.GetRecordType(new DssPath(SelectedPaths[i])) == RecordType.IrregularTimeSeries)
                    records.Add(r.GetTimeSeries(new DssPath(SelectedPaths[i])));
                if (r.GetRecordType(new DssPath(SelectedPaths[i])) == RecordType.PairedData)
                    records.Add(r.GetPairedData(new DssPath(SelectedPaths[i]).FullPath));
            }
            return records;
        }

        private void ExportRecords(List<object> records, ExcelWriter ew)
        {
            for (int i = 0; i < records.Count; i++)
            {
                if (records[i] is TimeSeries)
                    ew.Write(records[i] as TimeSeries, SelectedSheets[i]);

                if (records[i] is PairedData)
                    ew.Write(records[i] as PairedData, SelectedSheets[i]);
            }
        }

        public void GetAllSheets()
        {
            var s = new List<string>();
            ExcelReader er = new ExcelReader(DataFilePath);
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

        public bool AreSelectedSheetsRowCountsUniform()
        {
            foreach (var sheet in SelectedSheets)
            {
                ExcelReader r = new ExcelReader(DataFilePath);
                if (r.SmallestColumnRowCount(sheet) != r.RowCount(sheet))
                    return false;
            }
            return true;
        }
    }
}
