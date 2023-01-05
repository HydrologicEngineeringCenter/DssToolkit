using Hec.Excel;
using SpreadsheetGear;
using System;
using System.Windows;
using System.Windows.Controls;

namespace DssExcel
{
  /// <summary>
  /// Interaction logic for TimeSeriesReviewView.xaml
  /// </summary>
  public partial class PairedDataReviewView : UserControl
  {
    MainViewModel vm;
    public PairedDataReviewView(MainViewModel vm)
    {
      InitializeComponent();
      this.Loaded += PairedDataReviewView_Loaded;
      this.vm = vm;
    }

    private void PairedDataReviewView_Loaded(object sender, RoutedEventArgs e)
    {
      try
      {
        var pd = vm.GetPairedData();
        ExcelPairedData.Write(ExcelView.ActiveWorksheet, pd);
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
