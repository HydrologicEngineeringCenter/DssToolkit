using Hec.Dss;
using Hec.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace DssExcel
{
  /// <summary>
  /// Interaction logic for App.xaml
  /// </summary>
  public partial class App : Application
  {
    private static string GetUsage()
    {
      return "Wrong number of agrguments provided.  Usage:" +
             "\nDssExcel.exe -import-xls-to-dss-ui file.xls file.dss" +
             "\nDssExcel.exe -export-dss-to-excel file.dss file.xls path1;path2;...";
    }

    void DssExcel_Startup(object sender, StartupEventArgs e)
    {
      // Application is running
      // Process command line args
      if( e.Args.Length==0)
      {
        MessageBox.Show(GetUsage() );
        Close();
      }

      if (e.Args.Length>=1 && e.Args[0] == "-import-xls-to-dss-ui")
      {
        if (GetUIFileNames(out string excelFileName, out string dssFileName))
        {
          MainWindow mainWindow = new MainWindow(excelFileName, dssFileName);
          mainWindow.Show();
          return;        }
      }
      else if(e.Args.Length > 3 && e.Args[0] == "-export-dss-to-excel" )
      {
        ExportDssToExcel(e.Args);
      }
      Close();
    }

    private void ExportDssToExcel(string[] args)
    {
      string dssFileName = args[1];
      string excelFileName = args[2];
      var paths = args[3].Split(';');

      using (DssReader dss = new DssReader(dssFileName))
      {
        var tsList = new List<TimeSeries>();
        var pdList = new List<PairedData>();
        foreach (string path in paths)
        {
          DssPath p = new DssPath(path);
          var type = dss.GetRecordType(p);
          if (type is RecordType.RegularTimeSeries 
            || type is RecordType.IrregularTimeSeries)
          {
            var ts = dss.GetTimeSeries(p);
            tsList.Add(ts);
          }
          else if (type is RecordType.PairedData)
          {
            var pd = dss.GetPairedData(p.FullPath);
            pdList.Add(pd);
          }
        }
        ExcelTimeSeries.Write(excelFileName,tsList.ToArray());
        ExcelPairedData.Write(excelFileName,pdList.ToArray());
      }
    }

    private bool GetUIFileNames(out string excelFileName, out string dssFileName)
    {
      string[] commandLineArgs = Environment.GetCommandLineArgs();
      if (commandLineArgs.Length == 4)
      {
        // Validate file paths from command line
        if (File.Exists(commandLineArgs[2]) )
        {
          excelFileName = commandLineArgs[2];
          dssFileName = commandLineArgs[3];
          return true;
        }
        else
        {
          MessageBox.Show("Invalid Excel input file provided: '" + commandLineArgs[2]+"'.  Does that file exist?");
          excelFileName = dssFileName = "";
          return false;
        }
      }

      if (commandLineArgs.Length == 2)
      {
        if (TryGetFileName("Select Excel file", "Excel Files (*.xls;*.xlsx;*.csv)|*.xls;*.xlsx;*.csv", out excelFileName) &&
            TryGetFileName("Select DSS file", "DSS Files (*.dss)|*.dss", out dssFileName))
        {
          return true;
        }
      }

      MessageBox.Show(GetUsage());
      excelFileName = dssFileName = "";
      return false;
    }

    private bool TryGetFileName(string title, string filter, out string fileName)
    {
      var dialog = new Microsoft.Win32.OpenFileDialog
      {
        Title = title,
        Filter = filter
      };

      var dlgResult = dialog.ShowDialog();
      if (dlgResult.HasValue && dlgResult.Value)
      {
        fileName = dialog.FileName;
        return true;
      }

      fileName = "";
      return false;
    }

    private void Close()
    {
      Current.Shutdown();
    }
  }
}

