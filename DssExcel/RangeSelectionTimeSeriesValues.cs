using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  internal class RangeSelectionTimeSeriesValues : RangeSelectionVM
  {
    public RangeSelectionTimeSeriesValues(MainViewModel vm) : base(vm)
    {
      Title = "Select time series values";
      Description = "select one or more ranges with numbers";
    }

    public override bool Validate(out string errorMessage)
    {

      if( !ExcelReader.TryGetValueArray2D(RangeSelection, out double[,] values, out errorMessage))
      {
        return false;
      }

      mainViewModel.TimeSeriesValues = values;
      return true;

    }
  }
}
