using Hec.Dss;
using Hec.Excel;
using System;
using System.Collections.Generic;

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
    public string XValuesLabel { get; set; }
    public double[,] YValues { get; set; }
    public string YValuesLabel { get; set; }
    public double[,] TimeSeriesValues { get; set; }


    public string[] TimeSeriesNames { get; set; }

    public string DssFileName { get; set; }

    string firstRangeText1 = "";
    public string FirstRangeText { get =>firstRangeText1;
      set { firstRangeText1 = value; OnPropertyChanged(); }
    }
    string secondRangeText1 = "";
    public string SecondRangeText { get => secondRangeText1;
      set { secondRangeText1 = value; OnPropertyChanged(); } }
    
    internal Excel ExcelReader { get; set; }

    private static double[] GetColumn(double[,] matrix,int columnIndex)
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

    internal PairedData GetPairedData()
    {
      var rval = new PairedData();
      rval.Path = new DssPath(A: "", B: System.IO.Path.GetFileNameWithoutExtension(DssFileName),
        C: XValuesLabel + "-" + YValuesLabel, D: "", E: "", F: "");
      rval.Ordinates = this.XValues;
      rval.Values = new List<double[]>(rval.Ordinates.Length);
      
      for (int i = 0;i < YValues.GetLength(0); i++)
      {
        double[] row = new double[YValues.GetLength(1)];
        for (int j = 0; j < row.Length; j++)
        {
          row[j] = YValues[i, j];
          rval.Values.Add(row);
        }
      }
      
      return rval;
    }


  }


}
