using SpreadsheetGear;
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
      
    }


    public IWorksheet WorkSheet
    {
      get { return ExcelView.ActiveWorksheet; }
    }
  }
}
