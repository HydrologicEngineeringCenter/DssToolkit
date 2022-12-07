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
    MainViewModel vm;
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
        vm = new MainViewModel();
        vm.ExcelFileName = args[1];
        vm.DssFileName = args[2];
        vm.ExcelReader = new ExcelReader(vm.ExcelFileName); 
        statusControl.DataContext= vm;
        mainPanel.Content = new ImportTypeView(new ImportTypeVM());
        BackButton.IsEnabled= false;

      }
    }

    private void NextButton_Click(object sender, RoutedEventArgs e)
    {
      
      if (vm.ImportState == ImportState.SelectTimeSeries)
      {

        if (mainPanel.Content is ImportTypeView)
        {
          mainPanel.Content = new SelectDateRange(vm);
        }
      }
      else
      {
        mainPanel.Content = null;
      }
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
      if (mainPanel.Content is SelectDateRange)
      {
        mainPanel.Content = new ImportTypeView(new ImportTypeVM());
      }
    }
  }
}
