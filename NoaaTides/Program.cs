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
    static async Task Main()
    {
      //}",True,False,False,False,8772471,"Freeport Harbor",28.935699,-95.294197,NWLON,"","{
     var table = await NoaaTidesTimeSeries.ReadTimeSeries("8772471", "water_level", DateTime.Now.AddDays(-45), DateTime.Now.Date);

      table.WriteXml(@"c:\temp\freeport.xml");
     Console.WriteLine(table);
            // NoaaCoopsClient client = new NoaaCoopsClient();
            //var table = await client.GetStations();
        }
  }
}
