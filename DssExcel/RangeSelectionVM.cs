using SpreadsheetGear;

namespace DssExcel
{
  public abstract class RangeSelectionVM: ValidationVM
  {
    public RangeSelectionVM(MainViewModel vm)
    {
      mainViewModel = vm;
    }
    

    public MainViewModel mainViewModel;
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

    internal Excel ExcelReader { get => mainViewModel.ExcelReader; }

    IRange currentSelection;
    public IRange RangeSelection { get =>currentSelection; 
      internal set { currentSelection = value; OnPropertyChanged(); }
    }

  }
}
