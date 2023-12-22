using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Web;

namespace CwmsData.Api
{
  internal class CwmsDataClient
  {
    string REQUEST_JSONV2 = "application/json;version=2";
    string REQUEST_JSONV0 = "application/json";

    string officeID;
    string apiUrl;
    string apiKey;
    bool trustHttpServer = false;

    public string OfficeID { get => officeID; set => officeID = value; }

    public double MissingValue { get; set; }
    public CwmsDataClient(string apiUrl, string officeID)
    {
      MissingValue = double.MinValue;
      this.apiUrl = apiUrl;
      this.OfficeID = officeID;
      apiKey = Environment.GetEnvironmentVariable("CDA_API_KEY");

      Uri uri = new Uri(apiUrl);
      if( IsPrivateIpAddress( uri.Host))
        trustHttpServer = true;

      if( apiKey == null)
      {
        throw new Exception("Error: The environment variable: 'CDA_API_KEY' is not set");
      }
    }

    /// <summary>
    /// curl -X 'DELETE' \
    /// 'https://cwms-data.test:8444/cwms-data/timeseries/Mount%20Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev?office=SPK&begin=2020-06-10T13%3A00%3A00&end=2025-06-10T13%3A00%3A00' \
    /// -H 'accept: */*' \
    /// -H 'Authorization: apikey sessionkey-for-testing'
    /// </summary>
    /// <param name="name"></param>
    /// <returns></returns>
    internal async Task<Boolean> DeleteTimeSeries(string name, DateTime begin, DateTime end)
    {
      string encodedLocation = HttpUtility.UrlEncode(name);
      string encodedOffice = HttpUtility.UrlEncode(OfficeID);
      string t1 = HttpUtility.UrlEncode(begin.ToString("O"));
      string t2 = HttpUtility.UrlEncode(end.ToString("O"));

      // Combine the base URL with encoded location and office parameters
      string uri = $"{apiUrl}/timeseries/{encodedLocation}?office={encodedOffice}"
        + $"&begin={t1}&end={t2}";

      using (HttpClient client = GetClient())
      {
        HttpMethod method = HttpMethod.Delete;
        HttpRequestMessage request = new HttpRequestMessage(method, uri);
        request.Headers.Add("accept", "*/*");
        HttpResponseMessage response = await client.SendAsync(request);

        if (response.IsSuccessStatusCode)
        {
          return true;
        }
        else
        {
          Console.WriteLine($"Error: {response.StatusCode} - {response.ReasonPhrase}");
          return false;
        }
      }
    }

    internal async Task SaveTimeSeries(SimpleTimeSeries ts)
    {
      string json = ts.ToJson(this.officeID);
      await SaveTimeSeries(json);
    }
    internal async Task SaveTimeSeries(string json)
    {

      using (var client = GetClient())
      {
        string url = $"{apiUrl}/timeseries";

        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Authorization = null;
        client.DefaultRequestHeaders.Add("accept", "*/*");
        client.DefaultRequestHeaders.Add("Authorization", "apikey sessionkey-for-testing");


        var content = new StringContent(json, Encoding.UTF8);
        content.Headers.Remove("Content-Type");
        content.Headers.Add("Content-Type", REQUEST_JSONV2);

        var response = await client.PostAsync(url, content);
        response.EnsureSuccessStatusCode();
        var responseContent = await response.Content.ReadAsStringAsync();
        Console.WriteLine(responseContent);


      }

    }

    /// <summary>
    /// curl -X 'DELETE' \
    ///'https://cwms-data.test:8444/cwms-data/locations/karltest?office=SPK' \
    /// -H 'accept: */*'
    /// </summary>
    /// <param name="name"></param>
    /// <param name="office"></param>    /// <returns></returns>
    public async Task<bool> DeleteLocation(string name)
    {

      string encodedLocation = HttpUtility.UrlEncode(name);
      string encodedOffice = HttpUtility.UrlEncode(OfficeID);

      // Combine the base URL with encoded location and office parameters
      string uri = $"{apiUrl}/locations/{encodedLocation}?office={encodedOffice}";

      using (HttpClient client = GetClient())
      {
        HttpMethod method = HttpMethod.Delete;
        HttpRequestMessage request = new HttpRequestMessage(method, uri);
        request.Headers.Add("accept", "*/*");
        HttpResponseMessage response = await client.SendAsync(request);

        if (response.IsSuccessStatusCode)
        {
          return true;
        }
        else
        {
          Console.WriteLine($"Error: {response.StatusCode} - {response.ReasonPhrase}");
          return false;
        }
      }
    }

      public async Task<Location[]> GetLocations(string office="")
    {
      /*
       * curl -X 'GET' \
        'https://cwms-data.usace.army.mil/cwms-data/locations?office=SPK' \
         -H 'accept: application/json'
       */
      string endpoint = this.apiUrl+"/locations";
      if( office != "")
      {
        endpoint = endpoint + "?office=" + Uri.EscapeDataString(office);
      }

      string jsonData = await Get(endpoint,REQUEST_JSONV0);
      //File.WriteAllText(@"C:\project\cda-notes\location-response.json",jsonData);
      //string jsonData = await Task.Run(() => File.ReadAllText(@"C:\project\cda-notes\location-response.json"));
      jsonData = jsonData.Replace("\r\n", "\\n").Replace("\r", "\\r");
      var doc = JsonDocument.Parse(jsonData);
      var root = doc.RootElement;

      var locations = root.GetProperty("locations").GetProperty("locations");
      int size = locations.GetArrayLength();
      var rval = new List<Location>();

      foreach (JsonElement location in locations.EnumerateArray())
      {
        var val = new Location();
        val.Name = GetStringProperty(location, new[] { "identity", "name" });
        val.OfficeId = GetStringProperty(location, new[] { "identity", "office" });
        val.Latitude = GetDoubleProperty(location, new[] { "geolocation", "latitude" });
        val.Longitude = GetDoubleProperty(location, new[] { "geolocation", "longitude" });
        val.TimezoneName = GetStringProperty(location, new[] { "political" , "timezone" });
        val.LocationKind = GetStringProperty(location, new[] { "classification", "location-kind" });
        val.Nation = GetStringProperty(location, new[] { "political", "nation" });
        val.HorizontalDatum = GetStringProperty(location, new[] { "geolocation", "horizontal-datum" });
        rval.Add(val);
      }

        return rval.ToArray();

    }

    private static double GetDoubleProperty(JsonElement e, string[] propertyNames)
    {
      var s = GetStringProperty(e, propertyNames);
      double.TryParse(s, out double rval);
      return rval;
    }
    private static string GetStringProperty(JsonElement e, string[] propertyNames)
    {
      if (propertyNames.Length == 0)
        return "";
      JsonElement prop = e.GetProperty(propertyNames[0]);
      for (int i = 1; i < propertyNames.Length; i++)
      {
        prop = prop.GetProperty(propertyNames[i]);
      }
      return prop.ToString();
    }

    private async Task<string> Get(string url, string requestHeader )
    {
      string rval = "";
      using (HttpClient client = GetClient())
      {
        client.DefaultRequestHeaders.Accept.Clear();
        //client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(requestHeader));

        var request = new HttpRequestMessage
        {
          Method = HttpMethod.Get,
          RequestUri = new Uri(url)
        };
        request.Headers.Add("Accept", requestHeader);

        HttpResponseMessage response = await client.SendAsync(request, HttpCompletionOption.ResponseContentRead);
        response.EnsureSuccessStatusCode();

        rval = await response.Content.ReadAsStringAsync();
      }
      return rval;
    }


    /*
     * {
  "begin": "2023-06-23T06:01:00+0000[UTC]",
  "end": "2023-06-24T06:01:00+0000[UTC]",
  "interval": "PT0S",
  "interval-offset": 0,
  "name": "Mount Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev",
  "office-id": "LRB",
  "page": "MTY4NzUwMTgwMDAwMHx8NDh8fDUwMA==",
  "page-size": 500,
  "time-zone": "US/Eastern",
  "total": 48,
  "units": "ft",
  "value-columns": [
    {
      "name": "date-time",
      "ordinal": 1,
      "datatype": "java.sql.Timestamp"
    },
    {
      "name": "value",
      "ordinal": 2,
      "datatype": "java.lang.Double"
    },
    {
      "name": "quality-code",
      "ordinal": 3,
      "datatype": "int"
    }
  ],
  "values": [
    [1687501800000, 587.03, 0],
    [1687503600000, 587.03, 0],
    [1687505400000, 587.03, 0],
    [1687507200000, 587.03, 0],
    [1687509000000, 587.03, 0],
    [1687510800000, 587.03, 0],
    [1687512600000, 587.02, 0],
    [1687514400000, 587.02, 0],
    [1687516200000, 587.02, 0],
    [1687518000000, 587.02, 0],
    [1687519800000, 587.02, 0],
    [1687521600000, 587.02, 0],
    [1687523400000, 587.01, 0],
    [1687525200000, 587.01, 0],
    [1687527000000, 587.01, 0],
    [1687528800000, 587.01, 0],
    [1687530600000, 587.01, 0],
    [1687532400000, 586.9999999999999, 0],
    [1687534200000, 586.9999999999999, 0],
    [1687536000000, 586.9999999999999, 0],
    [1687537800000, 586.9999999999999, 0],
    [1687539600000, 586.9999999999999, 0],
    [1687541400000, 586.9999999999999, 0],
    [1687543200000, 586.9899999999999, 0],
    [1687545000000, 586.9899999999999, 0],
    [1687546800000, 586.98, 0],
    [1687548600000, 586.98, 0],
    [1687550400000, 586.98, 0],
    [1687552200000, 586.98, 0],
    [1687554000000, 586.9699999999999, 0],
    [1687555800000, 586.9699999999999, 0],
    [1687557600000, 586.9699999999999, 0],
    [1687559400000, 586.9699999999999, 0],
    [1687561200000, 586.9699999999999, 0],
    [1687563000000, 586.9699999999999, 0],
    [1687564800000, 586.9699999999999, 0],
    [1687566600000, 586.9699999999999, 0],
    [1687568400000, 586.9699999999999, 0],
    [1687570200000, 586.9699999999999, 0],
    [1687572000000, 586.9699999999999, 0],
    [1687573800000, 586.9699999999999, 0],
    [1687575600000, 586.98, 0],
    [1687577400000, 586.98, 0],
    [1687579200000, 586.98, 0],
    [1687581000000, 586.9899999999999, 0],
    [1687582800000, 586.9899999999999, 0],
    [1687584600000, 586.9899999999999, 0],
    [1687586400000, 586.9999999999999, 0]
  ],
  "vertical-datum-info": {
    "office": "LRB",
    "unit": "ft",
    "location": "Mount Morris",
    "native-datum": "NGVD-29",
    "elevation": 0.0,
    "offsets": [
      {
        "estimate": true,
        "to-datum": "NAVD-88",
        "value": -0.5353
      }
    ]
  }
}

     */
    public async Task<SimpleTimeSeries> ReadTimeSeries(string name, DateTime firstTime, DateTime lastTime)
    {
      /*
       * curl -X 'GET' \
  'https://cwms-data.usace.army.mil/cwms-data/timeseries?name=Mount%20Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev&office=LRB&begin=2023-06-23T06%3A01%3A00&end=2023-06-24T06%3A01%3A00' \
  -H 'accept: application/json;version=2'
       */
      var begin = firstTime.ToString("O");
      var end = lastTime.ToString("O");

      string queryString = $"?name={Uri.EscapeDataString(name)}&office={Uri.EscapeDataString(OfficeID)}&begin={Uri.EscapeDataString(begin)}&end={Uri.EscapeDataString(end)}";
      string apiUrlWithQuery = this.apiUrl + "/timeseries" + queryString;

      string jsonData = await Get(apiUrlWithQuery, REQUEST_JSONV2);

      var doc = JsonDocument.Parse(jsonData);
      var root = doc.RootElement;

      SimpleTimeSeries rval = new SimpleTimeSeries();
      rval.Name = root.GetProperty("name").ToString();
      rval.Units = root.GetProperty("units").ToString().Trim();

      var values = root.GetProperty("values");
      var len = values.GetArrayLength();

      if (values.GetArrayLength() <= 0)
      {
        Console.WriteLine("Warning: no time-series data found " + queryString);
        return rval;
      }

      DateTime t;
      foreach (JsonElement row in values.EnumerateArray())
      {
        long timestamp = row[0].GetInt64();

        var s = row[1].ToString();

        if (! double.TryParse(s,out double d))
             d = MissingValue;
        
        t = DateTimeOffset.FromUnixTimeMilliseconds(timestamp).DateTime;
        
        rval.Points.Add((t, d));
      }

      return rval;

    }

    

    public async Task<bool> SaveLocation(Location loc)
    {
      //      curl - X 'POST' \
      //  'https://cwms-data.localhost:8444/cwms-data/locations' \
      //  -H 'accept: */*' \
      //  -H 'Authorization: apikey <key-here>' \
      //  -H 'Content-Type: application/json' \
      //  -d '{
      //  "office-id": "SPK",
      //  "name": "karltest7",
      //  "latitude": 0,
      //  "longitude": 0,

      //  "location-kind": "SITE",
      //  "nation": "US",
      //  "horizontal-datum": "NAD83"
      //}'

      string url = this.apiUrl + "/locations";
      Console.WriteLine("Caling POST "+url);
      
      var json = loc.ToJson();
      Console.WriteLine("with Data:"+json);

      using (var client = GetClient())
      {
        
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));

        using (var content = new StringContent(json, Encoding.UTF8, "application/json"))
        {
          var response = await client.PostAsync(url, content);
          response.EnsureSuccessStatusCode();
          var responseContent = await response.Content.ReadAsStringAsync();
          Console.WriteLine(responseContent);
        }

        return true;
      }
    }

    private HttpClient GetClient()
    {
      var handler = new HttpClientHandler();
      if (trustHttpServer)
      {
        handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) =>
        {
          return true;
        };
        
      }

      var client = new HttpClient(handler);
      
      
      if (!string.IsNullOrWhiteSpace(apiKey))
      {
        client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("apikey", apiKey);
      }
      return client;
    }

    public static bool IsPrivateIpAddress(string hostName)
    {
      var hostEntry = Dns.GetHostEntry(hostName);
      foreach (var ip in hostEntry.AddressList)
      {
        if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
        {
          var bytes = ip.GetAddressBytes();

          switch (bytes[0])
          {
            case 10:
              return true;
            case 172:
              return bytes[1] < 32 && bytes[1] >= 16;
            case 192:
              return bytes[1] == 168;
            default:
              return false;
          }
        }
      }

      return false;
    }

  }
}
