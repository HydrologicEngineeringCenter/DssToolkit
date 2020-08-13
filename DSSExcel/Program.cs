using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;

namespace DSSExcelPlugin
{
    class Program
    {
        static void Main(string[] args)
        {
            DSSExcelLicensing licensing = new DSSExcelLicensing();
            licensing.SetPersonalLicense();

            // args layout: [dssFile] [excelFile] [action] [actionDestination]
            if (args.Length != 4)
                return;

            DSSExcel de = new DSSExcel(args[1]);
            if (args[2] == "import")
                de.Import(args[3], )

        }
    }
}
