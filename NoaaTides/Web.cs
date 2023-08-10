using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace NoaaTides
{
  internal class Web
  {
    internal async static Task<string> GetPage(string url)
    {
      using HttpClient client = new HttpClient();
      string rval = "";
      try
      {
        HttpResponseMessage response = await client.GetAsync(url);
        if (response.IsSuccessStatusCode)
        {
          rval = await response.Content.ReadAsStringAsync();
        }
        else
        {
          Console.WriteLine($"Failed to retrieve the web page. Status code: {response.StatusCode}");
        }
      }
      catch (Exception ex)
      {
        Console.WriteLine($"An error occurred reading: {url} {ex.Message}");
      }
      return rval;
    }
  }
}
