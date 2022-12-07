using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  public enum RangeSelectionType
  {
   TimeSeriesDateTimeColumn,
   TimeSeriesMultiColumn,
   PairedDataXColumn,
   PairedDataYColumns,
  }
  public class RangeSelectionVM:BaseVM
  {
    string _title;
    public string Title { get => _title; 
      set {  _title = value; OnPropertyChanged(); }
    }
    string _description;
    
    public string Description
    {
      get => _description;
      set { _description = value; OnPropertyChanged(); }
    }

    internal ExcelReader ExcelReader { get; set; }

    public RangeSelectionVM(RangeSelectionType rangeSelectionType, ExcelReader excel)
    {
      this.ExcelReader = excel;
      if(rangeSelectionType == RangeSelectionType.TimeSeriesDateTimeColumn)
      {
        Title = "Select Date/ Time Range: ";
        Description = "select rows in a single column. The date and time must be in the same column";
      }
      else if (rangeSelectionType == RangeSelectionType.TimeSeriesMultiColumn)
      {
        Title = "Select time series values";
        Description = "select one or more ranges with numbers";
      }
      else if( rangeSelectionType == RangeSelectionType.PairedDataXColumn)
      {
        Title = "Select X ordinate values";
        Description = "select values in a single column.)";
      }
      else if (rangeSelectionType == RangeSelectionType.PairedDataXColumn)
      {
        Title = "Select Y values";
        Description = "select values in a one or more columns. ";
      }
       else
      {
        Title = "";
        Description = "";
      }


    }
  }
}
