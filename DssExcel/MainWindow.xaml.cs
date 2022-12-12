using System;
using System.Collections.Generic;
using System.Windows;

namespace DssExcel
{
  public partial class MainWindow : Window
  {
    MainViewModel model;
    List<NavagationItem> timeSeriesControls = new List<NavagationItem>();
    int uiIndex = -1;
    public MainWindow()
    {
      InitializeComponent();

      GetFileNames(out string excelFileName, out string dssFileName);

      model = new MainViewModel();
      model.ExcelFileName = excelFileName;
      model.DssFileName = dssFileName;
      model.ExcelReader = new ExcelReader(model.ExcelFileName);
      statusControl.DataContext = model;
      CreateTimeSeriesNavagation();
      CreatePairedDataNavagation();

      NextButton_Click(this, new RoutedEventArgs());

    }

    private void GetFileNames(out string excelFileName, out string dssFileName)
    {
      excelFileName = "";
      dssFileName = "";
      string[] args = Environment.GetCommandLineArgs();
      if (args.Length == 3)
      {
        excelFileName = args[1];
        dssFileName = args[2];
      }
        else if (args.Length == 1)
      {// no args, prompt for filenames
        var dialog = new Microsoft.Win32.OpenFileDialog();
        dialog.Title = "Select Excel file";
        dialog.DefaultExt = ".xls";  
        dialog.Filter = "Excel Files (.xls)|*.xls";  
        var dlgResult = dialog.ShowDialog();
        if (dlgResult.HasValue && dlgResult.Value )
        {
         excelFileName = dialog.FileName;
          dialog.Title = "Select DSS file";
          dialog.DefaultExt = ".dss"; 
          dialog.Filter = "DSS Files (.dss)|*.dss"; 
          dlgResult = dialog.ShowDialog();
          if (dlgResult.HasValue && dlgResult.Value)
          {
            dssFileName = dialog.FileName;
          }
        }
        else
        {
          Close();
        }
      }
      else if (args.Length != 3)
      {
        excelFileName = "";
        dssFileName = "";
        MessageBox.Show("Wrong number of agrguments provided.  Usage:  DssExcel.exe file.xls file.dss");
        Close();
      }
    }

    private void CreatePairedDataNavagation()
    {
      
    }

    private void CreateTimeSeriesNavagation()
    {
      timeSeriesControls.Add(new NavagationItem
      {
        ViewModel = null,
        UserControl = new ImportTypeView(model.ImportTypeVM),
        BackEnabled = false,
        NextEnabled = true,
      });

      RangeSelectionVM vm = new RangeSelectionDatesVM(model);
      timeSeriesControls.Add(new NavagationItem
      {
        ViewModel = vm,
        UserControl = new RangeSelectionView(vm),
        BackEnabled = true,
        NextEnabled = true,
      });

      vm = new RangeSelectionTimeSeriesValues(model);
      timeSeriesControls.Add(new NavagationItem
      {
        ViewModel= vm,
        UserControl = new RangeSelectionView(vm),
        BackEnabled = true,
        NextEnabled = true,
      });

      
      timeSeriesControls.Add(new NavagationItem
      {
        ViewModel = null,
        UserControl = new TimeSeriesReviewView(model),
        BackEnabled = true,
        NextEnabled = false,
      });
    }

    private void NextButton_Click(object sender, RoutedEventArgs e)
    {

      model.ExcelReader.Workbook.WorkbookSet.GetLock();
      try
      {
        if (model.ImportType == ImportType.TimeSeries)
        {
          if (uiIndex >= 0) // initial uiIndex =-1
          {
            NavagationItem na = timeSeriesControls[uiIndex];
            var vm = na.ViewModel;
            if (vm != null && !vm.Validate(out string msg))
            {
              MessageBox.Show(msg);
              return;
            }
          }

          NavagationItem n = timeSeriesControls[++uiIndex];
          mainPanel.Content = n.UserControl;
          BackButton.IsEnabled = n.BackEnabled;
          NextButton.IsEnabled = n.NextEnabled;
        }
        else if (model.ImportType == ImportType.PairedData)
        {
          mainPanel.Content = null; // TO DO...
        }
        else
        {
          mainPanel.Content = null;
        }
      }
      finally{
        model.ExcelReader.Workbook.WorkbookSet.ReleaseLock();
      }
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
      if (model.ImportType == ImportType.TimeSeries)
      {
        NavagationItem n = timeSeriesControls[--uiIndex];
        mainPanel.Content = n.UserControl;
        BackButton.IsEnabled = n.BackEnabled;
        NextButton.IsEnabled = n.NextEnabled;
      }

    }
     
  }
}
