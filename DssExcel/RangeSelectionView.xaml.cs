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
  public partial class RangeSelectionView : UserControl
  {
    public RangeSelectionView(RangeSelectionVM vm)
    {
      InitializeComponent();
      this.DataContext = vm;
      this.ExcelView.ActiveWorkbook = vm.ExcelReader.Workbook;
      
    }

    DateTime[] DateTimes { get; }
    private void ExcelView_ActiveTabChanged(object sender, SpreadsheetGear.Windows.Controls.ActiveTabChangedEventArgs e)
    {

    }
  }
}
