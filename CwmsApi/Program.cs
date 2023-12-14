using CwmsApi;
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

      string apiUrl = "https://cwms-data.usace.army.mil/cwms-data";
      string officeID = "LRB";
      CwmsDataClient api = new CwmsDataClient(apiUrl, officeID);

      var x = await api.GetLocations("SPK");
      foreach (var location in x)
      {
        Console.WriteLine(location.Name);
      }
       
      //var x = await CwmsDataClient.PostLocation(CreateLocation("Test1"));


      string name = "Mount Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev";
      var begin = DateTime.Parse("2023-06-23T06:01:00");
      var end = DateTime.Parse("2023-06-24T06:01:00");

      var s = await api.GetTimeSeries( name, begin, end);
      s.WriteToConsole();


    

      }

    private static Location CreateLocation(string name, string locType ="SITE", string locKind = "PROJECT")
    {
      var location = new Location()
      {
        Name = name,
        TimezoneName = "US/Eastern",
        LocationType = "Dam",
        LocationKind = "PROJECT"
      };
      return location;
    }
  }
}
