using System;
using System.Threading.Tasks;

namespace CwmsData.Api
{
  internal static class Program
    {

    [STAThread]
    static async Task Main(string[] args)
    {
      string remoteApiUrl = "https://cwms-data.usace.army.mil/cwms-data";
      string localApiUrl = "https://cwms-data.test:8444/cwms-data";

      CwmsDataClient localAPI = new CwmsDataClient(localApiUrl, "SPK");
      CwmsDataClient remoteAPI = new CwmsDataClient(remoteApiUrl, "LRB");

      string name = "Mount Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev";

      Location mountMorris = new Location
      {
        OfficeId = localAPI.OfficeID,
        Name = "Mount Morris",
        Latitude = 0,
        Longitude = 0,
        TimezoneName = "UTC",
        LocationKind = "SITE",
        Nation = "US",
        HorizontalDatum = "NAD83"
      };

      await localAPI.SaveLocation(mountMorris);

      var t1 = DateTime.Parse("2023-06-23T06:01:00");
      var t2 = DateTime.Parse("2023-06-24T06:01:00");

      //await localAPI.DeleteTimeSeries(name, t1.AddYears(-100), t2.AddYears(10));
     
      Console.WriteLine($"Reading: {name}");
      var ts = await remoteAPI.ReadTimeSeries(name,t1, t2);
      ts.WriteToConsole();
      Console.WriteLine("Saving to Local Source");
      ts.Name = ts.Name + "-test";
      await localAPI.SaveTimeSeries(ts);

      Console.WriteLine("Reading back from  local source");
      ts = await localAPI.ReadTimeSeries(ts.Name, t1, t2);
      ts.WriteToConsole();
    }
  }
}
