using NoaaTides;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Media;
using System.Xml.Schema;
using Tools;

namespace TidesAndCurrents
{
  internal class NoaaMetaData
  {

    /// <summary>
    /// 
    /// </summary>
    /// <param name="stationType">waterlevels, waterlevelsandmet, airgap, datums, supersededdatums,
    /// benchmarks, supersededbenchmarks, historicwl, met, harcon, tidepredictions, currentpredictions, currents,
    /// historiccurrents, surveycurrents, cond, watertemp, physocean, tcoon, visibility, 1minute, historicmet, 
    /// historicphysocean, highwater, lowwater, hightideflooding, ofs</param>
    /// <returns></returns>
    public async Task<DataTable> GetStations(string stationType="waterlevels" )
    {
      if (!Regex.IsMatch(stationType, "^[\\w]{3,30}$"))
      {
        await Console.Out.WriteLineAsync("Error: invalid stationType '" + stationType + "'");
        return new DataTable();
      }

      string localCache = Path.Combine(Path.GetTempPath(), "tidesandcurrent_" + stationType + ".json");

      string url = "https://api.tidesandcurrents.noaa.gov/mdapi/prod/webapi/stations.json?units=english&type="+stationType;
      string json = "";
      JsonDocument doc;
      if (File.Exists(localCache))
      {
        json = File.ReadAllText(localCache);
      }
      else
      {
        json = await Web.GetPage(url);
        File.WriteAllText(localCache, json);
      }
      doc = JsonDocument.Parse(json);

      var table = GetTableFromJson(doc);
      return table;
    }
    private static DataTable GetTableFromJson(JsonDocument doc)
    {
      var rval = new DataTable();
      var root = doc.RootElement;

      if (root.TryGetProperty("stations", out JsonElement stations))
      {
        Console.WriteLine(stations.GetArrayLength());
        foreach (JsonElement station in stations.EnumerateArray())
        {
          var props = GetProperties(station);
          AddRowToTable(rval, props);
        }
      }


      return rval;
    }

    private static void AddRowToTable(DataTable table, Dictionary<string, string> props)
    {
      var row = table.NewRow();
      foreach (var item in props)
      {
        if (table.Columns.IndexOf(item.Key) < 0)
        {
          table.Columns.Add(item.Key, typeof(string));
        }

        if (IsImportantProperty(item))
            row[item.Key] = item.Value;
        
      }
      table.Rows.Add(row);
    }

    private static bool IsImportantProperty(KeyValuePair<string, string> item)
    {
      return !(item.Value.Contains("https://") || item.Value.Contains(","));
    }

    static Dictionary<string,string> GetProperties(JsonElement e)
    {
      var rval = new Dictionary<string,string>();
      
      foreach (JsonProperty item in e.EnumerateObject())
      {
        rval.Add(item.Name,item.Value.ToString());
      }
      return rval;
    }

  }
}

