using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using Hec.Dss;
using Hec.Dss.Excel;

namespace DSSExcel
{
    class Program
    {
        public class Options
        {
            [Option('c', "command", Required = true, HelpText = "The command for either importing or exporting DSS data and Excel data.")]
            public string Command { get; set; }

            [Option('d', "dss-file", Required = true, HelpText = "The source file used for exporting or importing from or to the destination file.")]
            public string DssFile { get; set; }

            [Option('e', "excel-file", Required = true, HelpText = "The destination file where the source file will export or import data.")]
            public string ExcelFile { get; set; }

            [Option('s', "excel-sheet", Required = true, HelpText = "The sheet in excel file used for importing or exporting data.")]
            public string Sheet { get; set; }

            [Option('p', "path", Required = false, HelpText = "Path of DSS Record in the form of '/a/b/c/d/e/f/'. (Required if exporting DSS data into excel)")]
            public string Path { get; set; }
        }

        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed(Run)
                .WithNotParsed(HandleParseError);
        }

        private static void HandleParseError(IEnumerable<Error> errs)
        {
            if (errs.IsVersion())
            {
                Console.WriteLine("Version Request");
                return;
            }

            if (errs.IsHelp())
            {
                Console.WriteLine("Help Request");
                return;
            }
            Console.WriteLine("Parser Fail");
        }

        private static void Run(Options opts)
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicense();
            if (opts.Command == "import")
            {
                if (!File.Exists(opts.ExcelFile))
                    throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");

                ExcelReader er = new ExcelReader(opts.ExcelFile);
                using (DssWriter w = new DssWriter(opts.DssFile))
                {
                    var t = er.CheckType(opts.Sheet);
                    if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                        w.Write(er.Read(opts.Sheet) as TimeSeries);
                    else if (t is RecordType.PairedData)
                        w.Write(er.Read(opts.Sheet) as PairedData);
                    
                }
            }
            else if (opts.Command == "export")
            {
                if (opts.Path == null)
                {
                    Console.WriteLine("DSS record path is needed for exporting data.");
                    return;
                }

                if (!File.Exists(opts.DssFile))
                    throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");

                using (DssReader r = new DssReader(opts.DssFile))
                {
                    object record;
                    ExcelWriter ew = new ExcelWriter(opts.ExcelFile);
                    DssPath path = new DssPath(opts.Path);
                    var type = r.GetRecordType(path);
                    if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                    {
                        record = r.GetTimeSeries(path);
                        ew.Write(record as TimeSeries, opts.Sheet);
                    }
                    else if (type is RecordType.PairedData)
                    {
                        record = r.GetPairedData(path.FullPath);
                        ew.Write(record as PairedData, opts.Sheet);
                    }


                }
            }
        }


    }
}
