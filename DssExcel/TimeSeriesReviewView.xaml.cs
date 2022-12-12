using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

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
      this.IsVisibleChanged += TimeSeriesReviewView_IsVisibleChanged;
      this.vm = vm;
    }

    private void TimeSeriesReviewView_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
      var etsc = new ExcelTimeSeriesCollectionVM(ExcelView.ActiveWorksheet);
      var locations = new string[vm.TimeSeriesNames.Length];
      for (int i = 0; i < vm.TimeSeriesNames.Length; i++)
        locations[i] = System.IO.Path.GetFileNameWithoutExtension(vm.ExcelFileName);

      etsc.Read(vm.DateTimes, vm.TimeSeriesValues, vm.TimeSeriesNames, locations);
    }
  }
}
