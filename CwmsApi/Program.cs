using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CwmsData.Api
{
    internal static class Program
    {

      [STAThread]
      static async Task Main(string[] args)
      {
      string name = "Mount Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev";
      string office = "LRB";
      var begin = DateTime.Parse("2023-06-23T06:01:00");
      var end = DateTime.Parse("2023-06-24T06:01:00");

      var s = await CwmsDataClient.GetTimeSeries(office, name, begin, end);
        s.WriteToConsole();
      //Console.WriteLine(s);

      }

   }
}
