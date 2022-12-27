using Hec.Dss;
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
    public ImportTypeVM ImportTypeVM { get; set; }

    public MainViewModel(ImportTypeVM importTypeVM)
    {
      this.ImportTypeVM = importTypeVM;
    }
    public string ExcelFileName { get; set; }

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

    private double[] GetColumn(double[,] matrix,int columnIndex)
    {
      var rval = new double[matrix.GetLength(0)];
      for (int i = 0; i < matrix.GetLength(0); i++)
      {
        rval[i] = matrix[i, columnIndex];
      }
      return rval;
    }
    public TimeSeries[] GetTimeSeries()
    {
      var rval = new List<TimeSeries>();
      for (int i = 0; i < TimeSeriesNames.Length; i++)
      {
        var ts = new TimeSeries();
        ts.Path = new DssPath(A:"",B:System.IO.Path.GetFileNameWithoutExtension(ExcelFileName), C:TimeSeriesNames[i],D:"",E:"", F:"xls-import");
        ts.Times = DateTimes;
        ts.Values = GetColumn(TimeSeriesValues,i);
        rval.Add(ts);
      }

      return rval.ToArray();
    }

  }


}
