using System;
using System.IO;
using System.Data;
using System.Windows.Media;

namespace NoaaTides
{
  internal class CsvFile : DataTable
  {
    public CsvFile(string fileName)
    {
      string[] lines = File.ReadAllLines(fileName);
      Parse(lines);
    }

    CsvFile()
    {
    }

    private void Parse(string[] lines)
    {
      var columnNames = lines[0].Split(',');
      for (int c = 0; c < columnNames.Length; c++)
      {
        Columns.Add(columnNames[c], typeof(String));
      }

      for (int i = 1; i < lines.Length; i++)
      {
        var line = lines[i];
        if (line.Trim() == "")
          continue;
        var tokens = line.Split(',');
        Rows.Add(tokens);
      }
    }

    internal static CsvFile FromString(string content)
    {
      var lines = content.Split(new char[] { '\n' });
      CsvFile result = new CsvFile();
      result.Parse(lines);
      return result;
    }

    static void ExportDataTableToCSV(DataTable dataTable, string filePath)
    {
      using (StreamWriter writer = new StreamWriter(filePath))
      {
        // Write header
        for (int i = 0; i < dataTable.Columns.Count; i++)
        {
          writer.Write(dataTable.Columns[i]);
          if (i < dataTable.Columns.Count - 1)
            writer.Write(",");
        }
        writer.WriteLine();

        // Write data
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
