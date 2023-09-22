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
      var s = await CwmsDataClient.GetTimeSeries();
        s.WriteToConsole();
      //Console.WriteLine(s);

      }

   }
}
