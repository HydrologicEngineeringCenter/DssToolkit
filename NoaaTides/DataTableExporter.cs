using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NoaaTides
{
  internal class DataTableExporter
  {
    /// <summary>
    /// Writes DataTable to a Text file that represents time-series data
    /// ExtendedProperties of the DataTable are included at the beginning
    /// of the text file as meta-data
    /// </summary>
    /// <param name="dataTable"></param>
    /// <param name="filePath"></param>
    internal static void WriteToCsv(DataTable dataTable, string filePath)
    {
      using (StreamWriter writer = new StreamWriter(filePath))
      {
        WriteHeader(dataTable, writer);
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
          writer.Write(dataTable.Columns[i]);
          if (i < dataTable.Columns.Count - 1)
            writer.Write(",");
        }
        writer.WriteLine();

        foreach (DataRow row in dataTable.Rows)
        {
          for (int i = 0; i < dataTable.Columns.Count; i++)
          {
            writer.Write(row[i]);
            if (i < dataTable.Columns.Count - 1)
              writer.Write(",");
          }
          writer.WriteLine();
        }
      }
    }

    private static void WriteHeader(DataTable dataTable, StreamWriter writer)
    {
      if (dataTable.ExtendedProperties.Count > 0)
      {
        writer.WriteLine("# Begin Header");
        foreach (DictionaryEntry prop in dataTable.ExtendedProperties)
        {
          if(!string.IsNullOrWhiteSpace(prop.Value.ToString()))
             writer.WriteLine("# " + prop.Key + "=" + prop.Value);
        }
        writer.WriteLine("# End Header");
      }
    }
  }
}
