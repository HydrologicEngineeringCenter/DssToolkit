using SpreadsheetGear;
using System;
using System.Windows;
using System.Windows.Controls;

namespace DssExcel
{
  /// <summary>
  /// Interaction logic for TimeSeriesReviewView.xaml
  /// </summary>
  public partial class TimeSeriesReviewView : UserControl
  {
    MainViewModel vm;
    public TimeSeriesReviewView(MainViewModel vm)
    {
      InitializeComponent();
      this.Loaded += TimeSeriesReviewView_Loaded;
      this.vm = vm;
    }

    private void TimeSeriesReviewView_Loaded(object sender, RoutedEventArgs e)
    {
      try
      {
        var locations = new string[vm.TimeSeriesNames.Length];
        var versionTags = new string[vm.TimeSeriesNames.Length];
        for (int i = 0; i < vm.TimeSeriesNames.Length; i++)
        {
          locations[i] = System.IO.Path.GetFileNameWithoutExtension(vm.ExcelFileName);
          versionTags[i] = "xls-import";
        }

        ExcelTimeSeries.Write(ExcelView.ActiveWorksheet, vm.DateTimes, vm.TimeSeriesValues,
          vm.TimeSeriesNames, locations, versionTags);
      }
      catch (Exception ex)
      {
        MessageBox.Show(ex.Message);
      }
    }

    public IWorksheet WorkSheet
    {
      get { return ExcelView.ActiveWorksheet; }
    }
  }
}
