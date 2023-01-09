using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  public class ImportTypeVM:BaseVM
  {
    public ObservableCollection<ImportOptionVM> ImportTypes { get; set; }

    private ImportOptionVM _selectedOptionVM;
    public ImportOptionVM SelectedImportType {
      get => _selectedOptionVM;
      set { _selectedOptionVM = value; OnPropertyChanged(); }
    }

    public ImportTypeVM()
    {
      ImportTypes = new ObservableCollection<ImportOptionVM>();

      ImportTypes.Add(new ImportOptionVM
      {
        Name = "Time Series Data",
        Description = "Import time series records by selecting data ranges in an excel workbook\n"
                                   + "Time Series have two components: 1 (date/time) and 2) numerical values. \n"
                                   + "There may be multiple value columns if several time series exist in your excel file.",
        Type = ImportType.TimeSeries,
      });

      ImportTypes.Add(new ImportOptionVM
      {
        Name = "Paired Data",
        Description = "Paired Data is two columns of data, eg. {x,y} where x=independent values, y= dependent values\n"
                                   + "Your paried data may also have multiple dependent values {x,y1,y2,...} ",
        Type = ImportType.PairedData,
      });

      SelectedImportType = ImportTypes[0];
    }
  }
}
