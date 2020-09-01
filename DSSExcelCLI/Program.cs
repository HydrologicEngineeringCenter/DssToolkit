using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using Hec.Dss;
using Hec.Dss.Excel;

namespace DSSExcelCLI
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

            [Option('s', "excel-sheet", Required = false, HelpText = "The sheet in excel file used for importing or exporting data.", Separator = ',')]
            public IEnumerable<string> Sheets { get; set; }

            [Option('p', "path", Required = false, HelpText = "Path of DSS Record in the form of '/a/b/c/d/e/f/'.", Separator = ',')]
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
            if (opts.Command == "import")
            {
                VerifyImportArgs(opts);

                ExcelReader er = new ExcelReader(opts.ExcelFile);
                using (DssWriter w = new DssWriter(opts.DssFile))
                {
                    if (opts.Sheets.ToList<string>().Count == 0)
                    {
                        for (int i = 0; i < er.Count; i++)
                        {
                            var t = er.CheckType(i);
                            if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                                w.Write(er.Read(i) as TimeSeries);
                            else if (t is RecordType.PairedData)
                                w.Write(er.Read(i) as PairedData);
                        }
                    }
                    else
                    {
                        foreach (var sheet in opts.Sheets)
                        {
                            var t = er.CheckType(sheet);
                            if (t is RecordType.RegularTimeSeries || t is RecordType.IrregularTimeSeries)
                                w.Write(er.Read(sheet) as TimeSeries);
                            else if (t is RecordType.PairedData)
                                w.Write(er.Read(sheet) as PairedData);
                        }
                    }
                    
                }
            }
            else if (opts.Command == "export")
            {

                VerifyExportArgs(opts);

                using (DssReader r = new DssReader(opts.DssFile))
                {
                    object record;
                    ExcelWriter ew = new ExcelWriter(opts.ExcelFile);
                    if (opts.Sheets.ToList<string>().Count == 0)
                    {
                        for (int i = 0; i < opts.Paths.ToList<string>().Count; i++)
                        {
                            DssPath p = new DssPath(opts.Paths.ElementAt(i));
                            var type = r.GetRecordType(p);
                            if (type is RecordType.RegularTimeSeries || type is RecordType.IrregularTimeSeries)
                            {
                                record = r.GetTimeSeries(p);
                                ew.Write(record as TimeSeries, i);
                            }
                            else if (type is RecordType.PairedData)
                            {
                                record = r.GetPairedData(p.FullPath);
                                ew.Write(record as PairedData, i);
                            }
                        }
                    }
                    else
                    {
                        for (int i = 0; i < opts.Sheets.ToList<string>().Count; i++)
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

        private static void VerifyImportArgs(Options opts)
        {
            if (!File.Exists(opts.ExcelFile))
                throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
        }

        private static void VerifyExportArgs(Options opts)
        {
            if (opts.Paths == null)
            {
                Console.WriteLine("DSS record path is needed for exporting data.");
                Environment.Exit(1);
            }

            if (opts.Paths.ToList<string>().Count != opts.Sheets.ToList<string>().Count)
            {
                Console.WriteLine("The sheet and path counts are not equal.");
                Environment.Exit(1);
            }

            if (!File.Exists(opts.DssFile))
                throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
        }
    }
}
