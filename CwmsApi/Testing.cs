using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CwmsData.Api
{
  internal static class Testing
  {

    /// <summary>
    /// Creates sample location
    /// </summary>
    /// <param name="api"></param>
    /// <returns></returns>
    internal static async Task CreateLocation(CwmsDataClient api, string name )
    {

      Location location1 = new Location
      {
        OfficeId = api.OfficeID,
        Name =name,
        Latitude = 0,
        Longitude = 0,
        TimezoneName = "UTC",
        LocationKind = "SITE",
        Nation = "US",
        HorizontalDatum = "NAD83"
      };


      var res = await api.PostLocation(location1);
    }

    internal static async Task<SimpleTimeSeries> ReadTimeSeries(CwmsDataClient api, string name, DateTime t1, DateTime t2)
    {
      Console.WriteLine($"Reading: {name}");
      var s = await api.ReadTimeSeries(name, t1, t2);
      s.WriteToConsole();
      return s;
    }


    internal static async Task ListLocations(CwmsDataClient api)
    {
      var x = await api.GetLocations(api.OfficeID);
      foreach (var location in x)
      {
        Console.WriteLine($"'{location.Name}'  at office: '{location.OfficeId}'");
      }
    }

    internal static async Task DeleteLocation(CwmsDataClient api, string name)
    {
      Console.WriteLine($"Calling Delete: '{name}'    office: '{api.OfficeID}'");
      var status = await api.DeleteLocation(name);
    }
  }
}
