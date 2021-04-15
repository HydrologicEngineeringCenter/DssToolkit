using Hec.Dss;
using Hec.Dss.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSSExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            string DssFile = args[1];
            string ExcelFile = args[2];

            using (DssReader r = new DssReader(DssFile))
            {
                object record;
                ExcelWriter ew = new ExcelWriter(opts.ExcelFile);
                if (opts.Sheets.ToList().Count == 0)
                {
                    for (int i = 0; i < opts.Paths.ToList<string>().Count; i++)
                    {
                        string sheet = "import_" + ExcelTools.RandomString(5);
                        DssPath p = new DssPath(opts.Paths.ElementAt(i));
                        var type = r.GetRecordType(p);
                        if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                        {
                            record = r.GetTimeSeries(p);
                            ew.Write(record as TimeSeries, sheet);
                        }
                        else if (type is RecordType.PairedData)
                        {
                            record = r.GetPairedData(p.FullPath);
                            ew.Write(record as PairedData, sheet);
                        }
                    }
                }
                else
                {
                    for (int i = 0; i < opts.Sheets.ToList().Count; i++)
                    {
                        DssPath p = new DssPath(opts.Paths.ElementAt(i));
                        var type = r.GetRecordType(p);
                        if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                        {
                            record = r.GetTimeSeries(p);
                            ew.Write(record as TimeSeries, opts.Sheets.ElementAt(i));
                        }
                        else if (type is RecordType.PairedData)
                        {
                            record = r.GetPairedData(p.FullPath);
                            ew.Write(record as PairedData, opts.Sheets.ElementAt(i));
                        }
                    }
                }
            }
        }
    }
}
