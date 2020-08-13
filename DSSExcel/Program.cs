using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using Hec.Dss;

namespace DSSExcelPlugin
{
    class Program
    {
        static void Main(string[] args)
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicense();
            
            if (args[0] == "import")
            {
                if (!File.Exists(args[1]))
                {
                    throw new FileNotFoundException("Couldn't find Excel file to import data into DSS.");
                }

                DSSExcel de = new DSSExcel(args[1]);
                de.Import(args[2], 0);
            }
            else if (args[0] == "export")
            {
                if (!File.Exists(args[1]))
                {
                    throw new FileNotFoundException("Couldn't find DSS file to import data into Excel.");
                }

                using (DssReader r = new DssReader(args[1]))
                {
                    var record = r.GetTimeSeries(new DssPath(args[3]));
                    DSSExcel.Export(args[2], record);
                }
            }
                

        }
    }
}
