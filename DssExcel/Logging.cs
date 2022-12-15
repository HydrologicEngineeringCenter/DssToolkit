using System;

namespace DssExcel
{
  internal class Logging
  {
    internal static void WriteError(string errorMessage)
    {
      Console.WriteLine("Error:"+errorMessage);
    }
  }
}