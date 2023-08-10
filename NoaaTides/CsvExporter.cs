using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NoaaTides
{
  internal class CsvExporter
  {
    static void WriteToCsv(DataTable dataTable, string filePath)
    {
      using (StreamWriter writer = new StreamWriter(filePath))
      {
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
  }
}
