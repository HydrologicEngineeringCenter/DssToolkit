using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Text.Json;
namespace CwmsData.Api
{
  internal class CwmsDataClient
  {
    public static async Task<string> GetJsonData()
    {
      /*
       * curl -X 'GET' \
  'https://cwms-data.usace.army.mil/cwms-data/timeseries?name=Mount%20Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev&office=LRB&begin=2023-06-23T06%3A01%3A00&end=2023-06-24T06%3A01%3A00' \
  -H 'accept: application/json;version=2'
       */
      string apiUrl = "https://cwms-data.usace.army.mil/cwms-data/timeseries";
      string name = "Mount Morris.Elev.Inst.30Minutes.0.GOES-NGVD29-Rev";
      string office = "LRB";
      string begin = "2023-06-23T06:01:00";
      string end = "2023-06-24T06:01:00";

      string queryString = $"?name={Uri.EscapeDataString(name)}&office={Uri.EscapeDataString(office)}&begin={Uri.EscapeDataString(begin)}&end={Uri.EscapeDataString(end)}";
      string apiUrlWithQuery = apiUrl + queryString;

      using (HttpClient client = new HttpClient())
      {
        client.DefaultRequestHeaders.Accept.Clear();
        client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

        HttpResponseMessage response = await client.GetAsync(apiUrlWithQuery,HttpCompletionOption.ResponseContentRead);
        response.EnsureSuccessStatusCode();

        string jsonData = await response.Content.ReadAsStringAsync();

        var doc = JsonDocument.Parse(jsonData);
        var root = doc.RootElement;
        var ts = root.GetProperty("time-series").GetProperty("time-series");
        if( ts.GetArrayLength() != 1)
        {
          throw new Exception("array length was " + ts.GetArrayLength() + " expected 1");
        }
        ts = ts[0];
        if( ts.TryGetProperty("regular-interval-values",out JsonElement rtsv ))
        {
          var interval = rtsv.GetProperty("interval");
        }
        return jsonData;
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
    }
  }

}
