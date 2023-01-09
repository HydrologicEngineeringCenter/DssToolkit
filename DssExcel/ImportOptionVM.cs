using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  public class ImportOptionVM : BaseVM
  {
    ImportType _type;
    public ImportType Type
    {
      get => _type;
      set { _type = value; OnPropertyChanged(); }
    }

    string _name;
    public string Name
    {
      get => _name;
      set { _name = value; OnPropertyChanged(); }
    }

    string _description;
    public string Description
    {
      get => _description;
      set { _description = value; OnPropertyChanged(); }

      //ICommand _command;
    }
  }
}
