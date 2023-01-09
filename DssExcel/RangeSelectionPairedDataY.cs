using Hec.Excel;

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

      mainViewModel.SecondRangeText = "Y: " + Excel.RangeToString(RangeSelection);
      mainViewModel.YValuesLabel = Excel.RangeTitle(RangeSelection, "Y");
      if (Excel.TryGetValueArray2D(RangeSelection, out double[,] values, out errorMessage))
      {
        mainViewModel.YValues = values;
        return true;
      }
      else { return false; }
    }

  }
}
