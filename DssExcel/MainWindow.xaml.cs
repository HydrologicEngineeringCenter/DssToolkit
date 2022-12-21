using System;
using System.Collections.Generic;
using System.Windows;

namespace DssExcel
{
  public partial class MainWindow : Window
  {
    MainViewModel model;
    List<NavigationItem> timeSeriesNavigation = new List<NavigationItem>();
    List<NavigationItem> pairedDataNavigation = new List<NavigationItem>();
    int uiIndex = -1;
    public MainWindow()
    {
      InitializeComponent();

      GetFileNames(out string excelFileName, out string dssFileName);

      model = new MainViewModel();
      model.ExcelFileName = excelFileName;
      model.DssFileName = dssFileName;
      model.ExcelReader = new Excel(model.ExcelFileName);
      statusControl.DataContext = model;

      var rootNavigation = new NavigationItem
      {
        ViewModel = null,
        UserControl = new ImportTypeView(model.ImportTypeVM),
        BackEnabled = false,
        NextEnabled = true,
      };

      CreateTimeSeriesNavagation(rootNavigation);
      CreatePairedDataNavagation(rootNavigation);

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

    private void CreatePairedDataNavagation(NavigationItem rootNavigation)
    {
      pairedDataNavigation.Add(rootNavigation);
      
       RangeSelectionVM vm = new RangeSelectionPairedDataX(model);

      pairedDataNavigation.Add(new NavigationItem
      {
        ViewModel = vm,
        UserControl = new RangeSelectionView(vm),
        BackEnabled = true,
        NextEnabled = true,
      });


    }

    private void CreateTimeSeriesNavagation(NavigationItem rootNavigation)
    {
      timeSeriesNavigation.Add(rootNavigation);

      RangeSelectionVM vm = new RangeSelectionDatesVM(model);
      timeSeriesNavigation.Add(new NavigationItem
      {
        ViewModel = vm,
        UserControl = new RangeSelectionView(vm),
        BackEnabled = true,
        NextEnabled = true,
      });

      vm = new RangeSelectionTimeSeriesValues(model);
      timeSeriesNavigation.Add(new NavigationItem
      {
        ViewModel= vm,
        UserControl = new RangeSelectionView(vm),
        BackEnabled = true,
        NextEnabled = true,
      });

    
      var reviewControl = new TimeSeriesReviewView(model);
      var reviewVM = new TimeSeriesReviewVM(reviewControl.WorkSheet, model.DssFileName);
    timeSeriesNavigation.Add(new NavigationItem
    {
      ViewModel = reviewVM,
      UserControl = reviewControl,
      BackEnabled = true,
      NextEnabled = true,
      FinalStep = true,
    }); 
    }

    private void NextButton_Click(object sender, RoutedEventArgs e)
    {
      List<NavigationItem> navagationItems;
      if (model.ImportType == ImportType.TimeSeries)
      {
        navagationItems = timeSeriesNavigation;
      }
      else // paired data
      {
        navagationItems = pairedDataNavigation;
      }

        model.ExcelReader.Workbook.WorkbookSet.GetLock();
      try
      {
        if (uiIndex >= 0) // initial uiIndex =-1
        {
          NavigationItem na = navagationItems[uiIndex];
          var vm = na.ViewModel;
          if (vm != null && !vm.Validate(out string msg))
          {
            MessageBox.Show(msg);
            return;
          }
          if (na.FinalStep)
          {
            Close();
            return;
          }
        }

        NavigationItem n = navagationItems[++uiIndex];
        mainPanel.Content = n.UserControl;
        BackButton.IsEnabled = n.BackEnabled;
        NextButton.IsEnabled = n.NextEnabled;
       
      }
      finally{
        model.ExcelReader.Workbook.WorkbookSet.ReleaseLock();
      }
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
      if (model.ImportType == ImportType.TimeSeries)
      {
        NavigationItem n = timeSeriesNavigation[--uiIndex];
        mainPanel.Content = n.UserControl;
        BackButton.IsEnabled = n.BackEnabled;
        NextButton.IsEnabled = n.NextEnabled;
      }

    }
     
  }
}
