using System.Reflection.Metadata;
using Tools;
namespace Usgs
{
  internal class ReadUsgsPluginFile
  {
    /*
     * Read usgs plugin file and output text file suitable for importing to a DSS file
     * 
     * get lat/long and other meta data from USGS webservice,
     * 
      Station Name=UKIAH CA
      Stream Name=RUSSIAN R
      Station ID=11461000
      Version Name=USGS
      Latitude=391144
      Longitude=1231138
      Elevation=599.22
      Coord Datum=NAD27
      Available Peak Start=15Mar1912
      Available Daily Start=01Oct1911
      Available Daily End=26Dec2016

     */
    static void Main(string[] args)
    {
      Console.WriteLine("pathname,x,y,xyDatum,xyUnits,coordSys,coordID,timeZone,usgs");
      string station = "";
      TextFile f = new TextFile("Russian.usgs");
      do
      {
        station = f.GetNext("Station Name");
        var stream = f.GetNext("Stream Name");
        var usgs = f.GetNext("Station ID");
        var s = "/" + stream + "/" + station + "/STAGE//15MIN/USGS/,,,2,4,1,0,PST," + usgs;
        Console.WriteLine(s);

      } while (station != "");
    }
  }
}