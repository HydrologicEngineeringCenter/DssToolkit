using System;

namespace Hec.Excel
{
  internal class Logging
  {
    internal static void WriteError(string errorMessage)
    {
      Console.WriteLine("Error:"+errorMessage);
    }
  }
}