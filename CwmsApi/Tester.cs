using System;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace CwmsData.Api{
  class Tester
  {
    
    static async Task Maina(string[] args)
    {
      var handler = new HttpClientHandler();
        handler.ServerCertificateCustomValidationCallback = (message, cert, chain, errors) =>
        {
          return true;
        };

      var client = new HttpClient(handler);
      
      client.DefaultRequestHeaders.Add("accept", "*/*");
      client.DefaultRequestHeaders.Add("Authorization", "apikey sessionkey-for-testing");

      var json = @"{
            ""name"": ""karltest.Elev.Inst.30Minutes.0.test"",
            ""office-id"": ""SPK"",
            ""units"": ""feet"",
            ""values"": [
                [
                    123,
                    54.3,
                    0
                ]
            ]
        }";


      HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, 
        "https://cwms-data.test:8444/cwms-data/timeseries");
    string JSONV2 = "application/json;version=2";
      using (var content = new StringContent(json, Encoding.UTF8))
      {
        request.Content = content;
        request.Content.Headers.Remove("Content-Type");
        request.Content.Headers.Add("Content-Type", JSONV2);

        var response = await client.SendAsync(request);

        string result = response.Content.ReadAsStringAsync().Result;
        Console.WriteLine(result);
      }
      
    }
  }
}
