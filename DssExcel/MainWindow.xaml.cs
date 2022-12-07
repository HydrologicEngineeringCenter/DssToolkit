using System;
using System.Collections.Generic;
using System.Windows;

namespace DssExcel
{
  public partial class MainWindow : Window
  {
    MainViewModel mvm;
    List<NavagationItem> timeSeriesControls = new List<NavagationItem>();
    int uiIndex = -1;
    public MainWindow()
    {
      InitializeComponent();
      string[] args = Environment.GetCommandLineArgs();
      if (args.Length != 3)
      {
        MessageBox.Show("Wrong number of agrguments provided.  Usage:  DssExcel.exe file.xls file.dss");
        Close();
      }
      else
      {
        mvm = new MainViewModel();
        mvm.ExcelFileName = args[1];
        mvm.DssFileName = args[2];
        mvm.ExcelReader = new ExcelReader(mvm.ExcelFileName);
        statusControl.DataContext = mvm;
        CreateTimeSeriesNavagation();
        CreatePairedDataNavagation();

        NextButton_Click(this, new RoutedEventArgs());
        
      }
    }

    private void CreatePairedDataNavagation()
    {
      
    }

    private void CreateTimeSeriesNavagation()
    {
      timeSeriesControls.Add(new NavagationItem
      {
        UserControl = new ImportTypeView(mvm.ImportTypeVM),
        BackEnabled = false,
        NextEnabled = true,
      });

      timeSeriesControls.Add(new NavagationItem
      {
        UserControl = new RangeSelectionView(new RangeSelectionVM(RangeSelectionType.TimeSeriesDateTimeColumn, mvm.ExcelReader)),
        BackEnabled = true,
        NextEnabled = true,
      });

      timeSeriesControls.Add(new NavagationItem
      {
        UserControl = new RangeSelectionView(new RangeSelectionVM(RangeSelectionType.TimeSeriesMultiColumn, mvm.ExcelReader)),
        BackEnabled = true,
        NextEnabled = true,
      });


      timeSeriesControls.Add(new NavagationItem
      {
        UserControl = new TimeSeriesReviewView(),
        BackEnabled = true,
        NextEnabled = false,
      });
    }

    private void NextButton_Click(object sender, RoutedEventArgs e)
    {

      if (mvm.ImportType == ImportType.TimeSeries)
      {
        NavagationItem n = timeSeriesControls[++uiIndex];
        mainPanel.Content = n.UserControl;
        BackButton.IsEnabled = n.BackEnabled;
        NextButton.IsEnabled = n.NextEnabled;
      }
      else if(mvm.ImportType == ImportType.PairedData)
      {
        mainPanel.Content = null; // TO DO...
      }
      else
      {
        mainPanel.Content = null;
      }
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
      if (mvm.ImportType == ImportType.TimeSeries)
      {
        NavagationItem n = timeSeriesControls[--uiIndex];
        mainPanel.Content = n.UserControl;
        BackButton.IsEnabled = n.BackEnabled;
        NextButton.IsEnabled = n.NextEnabled;
      }

    }
     
  }
}
