using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;

namespace DssExcel
{
  public enum ImportType
  {
    TimeSeries,
    PairedData,
  }


  public class MainViewModel : BaseVM
  {
    public ImportTypeVM ImportTypeVM;

    public MainViewModel()
    {
      ImportTypeVM = new ImportTypeVM();
    }
    public string ExcelFileName { get; set; }

    public ImportType ImportType
    {
      get { return ImportTypeVM.SelectedImportType.Type;}
    }

    public DateTime[] DateTimes { get; set; } 

    public double[] XValues { get; set; }
    public double[,] YValues { get; set; }
    public double[,] TimeSeriesValues { get; set; }
    public string[] TimeSeriesNames { get; set; }

    public string DssFileName { get; set; }

    string dateRangeText1 = "";
    public string DateRangeText { get =>dateRangeText1;
      set { dateRangeText1 = value; OnPropertyChanged(); }
    }
    string valueRangeText1 = "";
    public string ValueRangeText { get => valueRangeText1;
      set { valueRangeText1 = value; OnPropertyChanged(); } }
    
    internal Excel ExcelReader { get; set; }
  }


}
