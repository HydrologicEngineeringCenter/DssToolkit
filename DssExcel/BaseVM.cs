using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  public class BaseVM : INotifyPropertyChanged
  {
    public event PropertyChangedEventHandler PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string name = "")
    {
      PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
  }
}
