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
        var series = vm.GetTimeSeries();
        ExcelTimeSeries.Write(ExcelView.ActiveWorksheet, series);
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
