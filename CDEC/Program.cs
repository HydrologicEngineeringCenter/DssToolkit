using System.Text.RegularExpressions;
using Tools;
namespace CDEC
{
  internal class Program
  {
    static async Task Main(string[] args)
    {

      //CdecMetaData.GetStation("CLV");
      //http://cdec.water.ca.gov/dynamicapp/req/CSVLegacyDataServlet?station_id=CLV&sensor_num=16&dur_code=E&start_date=2006-01-01&end_date=2022-06-06

      var response = await Web.GetPage("http://cdec.water.ca.gov/dynamicapp/staMeta?station_id=CLV");
      //Console.WriteLine(response);
      string lat="", lng="";
      Match m = Regex.Match(response, "Latitude</b></td><td>(?<lat>[0-9\\.]+)");
      if (m.Success)
        lat = m.Groups["lat"].Value;
      m = Regex.Match(response, "Longitude</b></td><td>(?<lng>[-0-9\\.]+)");
      if (m.Success)
        lng = m.Groups["lng"].Value;

      Console.WriteLine(lat + " " + lng);
      /*
       <h2>RUSSIAN RIVER AT CLOVERDALE</h2>
    <a href="/webgis/?appid=cdecstation&sta=CLV">Map</a> of surrounding area<p>
    <table border=1>
      <tr><td><b>Station ID</b></td><td>CLV</td><td><b>Elevation</b></td><td>107 ft</td></tr>
      <tr><td><b>River Basin</b></td><td>RUSSIAN R</td><td><b>County</b></td><td>MENDOCINO</td></tr>
      <tr><td><b>Hydrologic Area</b></td><td>NORTH COAST</td><td><b>Nearby City</b></td><td>CLOVERDALE</td></tr>
      <tr><td><b>Latitude</b></td><td>38.879349&#176</td><td><b>Longitude</b></td><td>-123.053612&#176</td></tr>
      <tr><td><b>Operator</b></td><td>CA Dept of Water Resources/DFM-Hydro-SMN</td><td><b>Maintenance</b></td><td>CA Dept of Water Resources/DFM-Hydro-SMN</td></tr>
    </table>
    <p>
       */
    }
  }
}