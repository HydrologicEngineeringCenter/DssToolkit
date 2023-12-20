using System;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

namespace CwmsData.Api { 
  class LoggingHandler : DelegatingHandler
  {
    protected override async Task<HttpResponseMessage> SendAsync(
        HttpRequestMessage request, CancellationToken cancellationToken)
    {
      // Log request headers before sending the request
      Console.WriteLine("Request Headers:");
      foreach (var header in request.Headers)
      {
        Console.WriteLine($"{header.Key}: {string.Join(", ", header.Value)}");
      }

      if (request.Content != null)
      {
        Console.WriteLine("Content Headers:");
        foreach (var header in request.Content.Headers)
        {
          Console.WriteLine($"{header.Key}: {string.Join(", ", header.Value)}");
        }
      }

      // Continue with the request
      var response = await base.SendAsync(request, cancellationToken);

      // Log response headers after receiving the response
      Console.WriteLine("Response Headers:");
      foreach (var header in response.Headers)
      {
        Console.WriteLine($"{header.Key}: {string.Join(", ", header.Value)}");
      }

      if (response.Content != null)
      {
        Console.WriteLine("Content Headers:");
        foreach (var header in response.Content.Headers)
        {
          Console.WriteLine($"{header.Key}: {string.Join(", ", header.Value)}");
        }
      }

      return response;
    }
  }

  
}