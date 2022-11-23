using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{

  public class DssExcelViewModel:INotifyPropertyChanged
  {
    private string[] importTypeNames = { "Time Series Data", "Paried Data" };
    private string[] importTypeDescription = { "Import time series records by selecting data ranges in an excel workbook\n"
                                   +"Time Series have two components: 1 (date/time) and 2) numerical values. \n"
                                   +"There may be multiple value columns if several time series exist in your excel file.",
                                   "Paired Data is two columns of data, eg. {x,y} where x=independent values, y= dependent values\n"
                                   +"Your paried data may also have multiple dependent values {x,y1,y2,...} " };

    public DssExcelViewModel()
    {
      ImportTypes = new List<string>();
      ImportTypes.AddRange(importTypeNames);
      //SelectedImportType = importTypeNames[0];
    }
    public string ExcelFileName { get; set; } 

    public string DssFileName { get; set; }

    public List<string> ImportTypes { get; set; }

    int selectedImportIndex;
    public int SelectedImportIndex
    {
      get { return selectedImportIndex; }
      set
      {
        selectedImportIndex = value;
        OnPropertyChanged("SelectedImportTypeDescription");
      }
    }
    public string SelectedImportTypeDescription
    {
      get
      {
        if (SelectedImportIndex >= 0)
          return importTypeDescription[SelectedImportIndex];
        return "hi";
      }
    }

    public event PropertyChangedEventHandler PropertyChanged;

    public void OnPropertyChanged(string name)
    {
      if (PropertyChanged == null)
        return;
      PropertyChanged(this, new PropertyChangedEventArgs(name));
    }
  }


}
