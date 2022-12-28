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
      mainViewModel.FirstRangeText = "X: " + Excel.RangeToString(RangeSelection);
      mainViewModel.XValuesLabel = Excel.RangeTitle(RangeSelection, "X");
      if (Excel.TryGetValueArray(RangeSelection, out double[] values, out errorMessage))
      {
        mainViewModel.XValues = values;
        return true;
      }
      else{ return false; }
    }
  }
}
