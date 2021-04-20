using Hec.Dss;
using Hec.Dss.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using System.IO;

namespace DSSExcel
{
    class Program
    {
        public class Options
        {
            [Option('d', "dss-file", Required = true, HelpText = "The source file used for exporting or importing from or to the destination file.")]
            public string DssFile { get; set; }

            [Option('e', "excel-file", Required = true, HelpText = "The destination file where the source file will export or import data.")]
            public string ExcelFile { get; set; }

            [Option('p', "paths", Required = true, HelpText = "Path of DSS Record in the form of '/a/b/c/d/e/f/'.", Separator = ';')]
            public IEnumerable<string> Paths { get; set; }
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
            VerifyExportArgs(opts);

            using (DssReader r = new DssReader(opts.DssFile))
            {
                object record;
                ExcelWriter ew = new ExcelWriter(opts.ExcelFile);
                
                for (int i = 0; i < opts.Paths.ToList<string>().Count; i++)
                {
                    string sheet;
                    while (true)
                    {
                        sheet = "dssvue_import_" + ExcelTools.RandomString(5);
                        if (!ew.SheetExists(sheet))
                            break;
                    }
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
            
        }

        private static void VerifyExportArgs(Options opts)
        {
            if (opts.Paths == null)
            {
                Console.WriteLine("DSS record path is needed for exporting data.");
                Environment.Exit(1);
            }

            if (!File.Exists(opts.DssFile))
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
        }
    }
}
