using SpreadsheetGear;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  public abstract class RangeSelectionVM:BaseVM
  {
    public RangeSelectionVM(MainViewModel vm)
    {
      mainViewModel = vm;
    }
    public abstract bool Validate(out string errorMessage);

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

    internal ExcelReader ExcelReader { get => mainViewModel.ExcelReader; }

    IRange currentSelection;
    public IRange RangeSelection { get =>currentSelection; 
      internal set { currentSelection = value; OnPropertyChanged(); }
    }

  }
}
