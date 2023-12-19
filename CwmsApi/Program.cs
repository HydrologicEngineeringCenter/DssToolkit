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


      await Testing.CreateLocation(localAPI, "karltest");
      await Testing.ListLocations(localAPI);
      await Testing.DeleteLocation(localAPI, "karltest");
      await Testing.ReadTimeSeries(remoteAPI, "Mount Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev",
        DateTime.Parse("2023-06-23T06:01:00"), DateTime.Parse("2023-06-24T06:01:00"));

      }
    
  }
}
