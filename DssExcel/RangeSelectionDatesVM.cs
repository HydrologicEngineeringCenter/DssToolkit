using Hec.Excel;
using System;

namespace DssExcel
{
  internal class RangeSelectionDatesVM : RangeSelectionVM
  {
    public RangeSelectionDatesVM(MainViewModel vm) : base(vm)
    {
      Title = "Select Date/ Time Range: ";
      Description = "select rows in a single column. The date and time must be in the same column";
    }

    public override bool Validate(out string errorMessage)
    {
      errorMessage = "";
      mainViewModel.FirstRangeText = "Dates: "+Excel.RangeToString(RangeSelection);
      
      if( Excel.TryGetDateArray(RangeSelection,out DateTime[] dates,out errorMessage))
      {
        mainViewModel.DateTimes = dates;
        return true;
      }
      else { return false; }
    }
  }
}
