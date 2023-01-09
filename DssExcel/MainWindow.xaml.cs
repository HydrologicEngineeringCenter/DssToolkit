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
    public MainWindow(string excelFileName, string dssFileName)
    {
      InitializeComponent();

      model = new MainViewModel(new ImportTypeVM());
      model.ExcelFileName = excelFileName;
      model.DssFileName = dssFileName;
      model.ExcelReader = new Excel(model.ExcelFileName);
      navigation = new NavigationCollection(model);
      statusControl.DataContext = model;
      
      NextButton_Click(this, new RoutedEventArgs());

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
