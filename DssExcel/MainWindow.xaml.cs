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
  public partial class MainWindow : Window
  {
    DssExcelViewModel vm;
    public MainWindow()
    {
      InitializeComponent();
      string[] args = Environment.GetCommandLineArgs();
      if (args.Length != 3)
      {
        MessageBox.Show("No agrguments provided.  Usage:  DssExcel.exe file.xls file.dss");
        Close();
      }
      else
      {
        vm = new DssExcelViewModel();
        vm.ExcelFileName = args[1];
        vm.DssFileName = args[2];

        mainPanel.Content = new SelectDataTypeView(vm);
      }
    }
  }
}
