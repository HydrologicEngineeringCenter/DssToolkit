using CwmsApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace CwmsData.Api
{
  internal class CwmsDataClient
  {
    string officeID;
    string apiUrl;
    string apiKey;
    string certificateFileName;
    string certificatePassword;
    public CwmsDataClient(string apiUrl, string officeID)
    {
      this.apiUrl = apiUrl;
      this.officeID = officeID;
      apiKey = Environment.GetEnvironmentVariable("CDA_API_KEY");
      certificateFileName = Environment.GetEnvironmentVariable("CDA_CERTIFICATE_FILENAME");
      certificatePassword = Environment.GetEnvironmentVariable("CDA_CERTIFICATE_PASSWORD");

      if( apiKey == null)
      {
        throw new Exception("Error: The environment variable: 'CDA_API_KEY' is not set");
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

      string jsonData = await Get(endpoint);
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

    private async Task<string> Get(string url)
    {
      string rval = "";
      using (HttpClient client = new HttpClient())
      {
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        HttpResponseMessage response = await client.GetAsync(url, HttpCompletionOption.ResponseContentRead);
        response.EnsureSuccessStatusCode();

        rval = await response.Content.ReadAsStringAsync();
      }
      return rval;
    }

    public async Task<SimpleTimeSeries> GetTimeSeries(string name, DateTime firstTime, DateTime lastTime)
    {
      /*
       * curl -X 'GET' \
  'https://cwms-data.usace.army.mil/cwms-data/timeseries?name=Mount%20Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev&office=LRB&begin=2023-06-23T06%3A01%3A00&end=2023-06-24T06%3A01%3A00' \
  -H 'accept: application/json;version=2'
       */
      var begin = firstTime.ToString("O");
      var end = lastTime.ToString("O");

      string queryString = $"?name={Uri.EscapeDataString(name)}&office={Uri.EscapeDataString(officeID)}&begin={Uri.EscapeDataString(begin)}&end={Uri.EscapeDataString(end)}";
      string apiUrlWithQuery = this.apiUrl + "/timeseries" + queryString;

        string jsonData = await Get(apiUrlWithQuery);

        var doc = JsonDocument.Parse(jsonData);
        var root = doc.RootElement;
        var ts = root.GetProperty("time-series").GetProperty("time-series");
        if (ts.GetArrayLength() != 1)
        {
          throw new Exception("array length was " + ts.GetArrayLength() + " expected 1");
        }
        ts = ts[0];
        //ValueKind = Object : "{"office":"LRB","name":"Mount Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev",
        ////   "alternate-names":["NY00468.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev"],
        /// "regular-interval-values":{"interval":"PT30M","unit":"ft NGVD29","segment-count":1,
        /// "segments":
        /// [{"first-time":"2023-06-23T06:30:00Z","last-time":"2023-06-24T06:00:00Z",
        ///  "value-count":48,"comment":"value, quality code","values":[[587.03,0],[587.03,0],[587.03,0],[587.03,0],[587.03,0],[587.03,0],[587.02,0],[587.02,0],[587.02,0],[587.02,0],[587.02,0],[587.02,0],[587.01,0],[587.01,0],[587.01,0],[587.01,0],[587.01,0],[587,0],[587,0],[587,0],[587,0],[587,0],[587,0],[586.99,0],[586.99,0],[586.98,0],[586.98,0],[586.98,0],[586.98,0],[586.97,0],[586.97,0],[586.97,0],[586.97,0],[586.97,0],[586.97,0],[586.97,0],[586.97,0],[586.97,0],[586.97,0],[586.97,0],[586.97,0],[586.98,0],[586.98,0],[586.98,0],[586.99,0],[586.99,0],[586.99,0],[587,0]]}
        /// ]}}"
        SimpleTimeSeries rval = new SimpleTimeSeries();
      if (ts.TryGetProperty("regular-interval-values", out JsonElement rtsv))
      {

        var interval = rtsv.GetProperty("interval").ToString();

        TimeSpan duration = System.Xml.XmlConvert.ToTimeSpan(interval);

        var units = rtsv.GetProperty("unit");
        var segmentCount = rtsv.GetProperty("segment-count");
        var segments = rtsv.GetProperty("segments");
        foreach (JsonElement segment in segments.EnumerateArray())
        {
          var first = segment.GetProperty("first-time").GetDateTime();
          var last = segment.GetProperty("last-time").GetDateTime();
          var t = firstTime;
          var values = segment.GetProperty("values");
          foreach (JsonElement value in values.EnumerateArray())
          {
            var str = value.EnumerateArray().First().ToString();
            if (double.TryParse(str, out double v))
            {
              rval.Points.Add((t, v));
            }
            t = t.Add(duration);
          }

        }

      }
        return rval;
        /*
         * {
  "time-series": {
    "query-info": {
      "time-of-query": "2023-06-26T23:31:13Z",
      "process-query": "PT0.473S",
      "format-output": "PT0.003S",
      "requested-start-time": "2023-06-23T06:01:00Z",
      "requested-end-time": "2023-06-24T06:01:00Z",
      "requested-format": "JSON",
      "requested-office": "LRB",
      "requested-items": [
        {
          "name": "Mount Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev",
          "unit": "EN",
          "datum": "NATIVE"
        }
      ],
      "total-time-series-retrieved": 1,
      "unique-time-series-retrieved": 1,
      "total-values-retrieved": 48,
      "unique-values-retrieved": 48
    },
    "quality-codes": {
      "comment": "The following quality codes are used in this dataset",
      "codes": [
        {
          "code": 0,
          "meaning": "Unscreened"
        }
      ]
    },
    "time-series": [
      {
        "office": "LRB",
        "name": "Mount Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev",
        "alternate-names": [
          "NY00468.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev"
        ],
        "regular-interval-values": {
          "interval": "PT30M",
          "unit": "ft NGVD29",
          "segment-count": 1,
          "segments": [
            {
              "first-time": "2023-06-23T06:30:00Z",
              "last-time": "2023-06-24T06:00:00Z",
              "value-count": 48,
              "comment": "value, quality code",
              "values": [
                [
                  587.03,
                  0
                ],
                [
                  587.03,
                  0
                ],
                [
                  587.03,
                  0
                ],
                [
                  587.03,
                  0
                ],
                [
                  587.03,
                  0
                ],
                [
                  587.03,
                  0
                ],
                [
                  587.02,
                  0
                ],
                [
                  587.02,
                  0
                ],
                [
                  587.02,
                  0
                ],
                [
                  587.02,
                  0
                ],
                [
                  587.02,
                  0
                ],
                [
                  587.02,
                  0
                ],
                [
                  587.01,
                  0
                ],
                [
                  587.01,
                  0
                ],
                [
                  587.01,
                  0
                ],
                [
                  587.01,
                  0
                ],
                [
                  587.01,
                  0
                ],
                [
                  587,
                  0
                ],
                [
                  587,
                  0
                ],
                [
                  587,
                  0
                ],
                [
                  587,
                  0
                ],
                [
                  587,
                  0
                ],
                [
                  587,
                  0
                ],
                [
                  586.99,
                  0
                ],
                [
                  586.99,
                  0
                ],
                [
                  586.98,
                  0
                ],
                [
                  586.98,
                  0
                ],
                [
                  586.98,
                  0
                ],
                [
                  586.98,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.97,
                  0
                ],
                [
                  586.98,
                  0
                ],
                [
                  586.98,
                  0
                ],
                [
                  586.98,
                  0
                ],
                [
                  586.99,
                  0
                ],
                [
                  586.99,
                  0
                ],
                [
                  586.99,
                  0
                ],
                [
                  587,
                  0
                ]
              ]
            }
          ]
        }
      }
    ]
  }
}
         */
      

    }





    public async Task<bool> PostLocation(Location loc)
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

      var json = loc.ToJson();

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
      if (certificateFileName != null && certificatePassword != null)
      {
        var certificate = new X509Certificate2(certificateFileName, certificatePassword);
        
        handler.ClientCertificates.Add(certificate);
        handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) =>
        {
          bool test = cert.GetCertHashString() == certificate.GetCertHashString();
          Console.WriteLine(test);
          return test;
        };
      }

      var client = new HttpClient(handler);
      client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("apikey", apiKey);
      return client;
    }

  }
}
