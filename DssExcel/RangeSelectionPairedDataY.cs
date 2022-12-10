using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  internal class RangeSelectionPairedDataY : RangeSelectionVM
  {
    public RangeSelectionPairedDataY(MainViewModel vm) : base(vm)
    {
      Title = "Select Y values";
      Description = "select values in a one or more columns. ";
    }

    public override bool Validate(out string errorMessage)
    {
      if (ExcelReader.TryGetValueArray2D(RangeSelection, out double[,] values, out errorMessage))
      {
        mainViewModel.YValues = values;
        return true;
      }
      else { return false; }
    }

  }
}
