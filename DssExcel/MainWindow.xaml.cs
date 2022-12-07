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
    MainViewModel mvm;
    ImportTypeVM importTypeVM;
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
        statusControl.DataContext= mvm;
        importTypeVM = new ImportTypeVM();
        mainPanel.Content = new ImportTypeView(importTypeVM);
        Enabling();
      }
    }

    private void NextButton_Click(object sender, RoutedEventArgs e)
    {
      Console.WriteLine(importTypeVM.SelectedImportType.Type.ToString());
      if (mainPanel.Content is ImportTypeView)
       {
         mainPanel.Content = new SelectDateRange(mvm);
       }
      else
      {
        mainPanel.Content = null;
      }
      Enabling();
    }

    private void BackButton_Click(object sender, RoutedEventArgs e)
    {
      if (mainPanel.Content is SelectDateRange)
      {
        mainPanel.Content = new ImportTypeView(importTypeVM);
      
      }
      Enabling();
    }

    private void Enabling()
    {
      BackButton.IsEnabled = !(mainPanel.Content is ImportTypeView);
    }
  }
}
