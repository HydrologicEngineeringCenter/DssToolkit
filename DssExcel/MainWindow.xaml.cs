using Hec.Excel;
using System;
using System.Collections.Generic;
using System.Windows;

namespace DssExcel
{
  public partial class MainWindow : Window
  {
    MainViewModel model;
    NavigationCollection navigation;

    int uiIndex = -1;
    public MainWindow()
    {
      InitializeComponent();

      GetFileNames(out string excelFileName, out string dssFileName);

      model = new MainViewModel(new ImportTypeVM());
      model.ExcelFileName = excelFileName;
      model.DssFileName = dssFileName;
      model.ExcelReader = new Excel(model.ExcelFileName);
      navigation = new NavigationCollection(model);
      statusControl.DataContext = model;
      
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


    private void NextButton_Click(object sender, RoutedEventArgs e)
    {
      NavigationClicked(true);
    }

    private void NavigationClicked(bool forward)
    {
      model.ExcelReader.Workbook.WorkbookSet.GetLock();
      try
      {
        if (uiIndex >= 0 && forward) // initial uiIndex =-1
        {
          NavigationItem na = navigation[uiIndex];
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
        // move to next screen...
        uiIndex += forward? 1 : -1;

        NavigationItem n = navigation[uiIndex];
        mainPanel.Content = n.UserControl;
        BackButton.IsEnabled = n.BackEnabled;
        NextButton.IsEnabled = n.NextEnabled;

      }
      finally
      {
        model.ExcelReader.Workbook.WorkbookSet.ReleaseLock();
      }
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
      NavigationClicked(false);
    }
     
  }
}
