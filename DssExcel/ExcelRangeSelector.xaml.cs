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
  /// Interaction logic for ExcelRangeSelector.xaml
  /// </summary>
  public partial class ExcelRangeSelector : UserControl
  {
    private ExcelReader excelReader1 = null;
    public ExcelRangeSelector(ExcelReader excelReader)
    {
      InitializeComponent();
      excelReader1 = excelReader; 
    }

    private void ExcelView_RangeSelectionChanged(object sender, SpreadsheetGear.Windows.Controls.RangeSelectionChangedEventArgs e)
    {
      //e.RangeSelection
    }
  }
}
