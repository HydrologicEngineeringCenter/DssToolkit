using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TidesAndCurrents;

namespace NoaaTides
{
  internal class Program
  {
    static async Task Main(string[] args)
    {
      //var m = new NoaaMetaData();
      //var stations = await m.GetStations();
      //DataTableExporter.WriteToCsv(stations, @"C:\project\DssToolkit\NoaaTides\stations.csv");
      

      var ts = await NoaaTidesTimeSeries.ReadTimeSeries("8772471", "water_level", DateTime.Now.AddDays(-45), DateTime.Now.Date);
      DataTableExporter.WriteToCsv(ts, @"c:\temp\ts.txt");
     Console.WriteLine(ts);


     }
  }
}
