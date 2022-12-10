using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  internal class RangeSelectionPairedDataX : RangeSelectionVM
  {
    public RangeSelectionPairedDataX(MainViewModel vm) : base(vm)
    {
      Title = "Select X ordinate values";
      Description = "select values in a single column.)";
    }

    public override bool Validate(out string errorMessage)
    {
      errorMessage = "";
      if (RangeSelection.ColumnCount != 1)
      {
        errorMessage = "Please select values in a single column";
        return false;
      }

      if (ExcelReader.TryGetValueArray(RangeSelection, out double[] values, out errorMessage))
      {
        mainViewModel.XValues = values;
        return true;
      }
      else{ return false; }
    }
  }
}
