using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CwmsApi
{
  internal class SimpleTimeSeries
  {
    public List<(DateTime Timestamp, double Value)> Points = new List<(DateTime, double)>();
    public Dictionary<string, string> Attributes = new();

    public void WriteToConsole()
    {
      foreach (var p in Points)
      {
        Console.WriteLine(p.Timestamp.ToString() + "," + p.Value);
      }
    }
  }
}
