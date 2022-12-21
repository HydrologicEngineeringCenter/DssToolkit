using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace DssExcel
{
  public  class NavigationItem
  {
    public bool FinalStep { get; set; }

    public ValidationVM ViewModel { get; set; }
    public UserControl UserControl { get; set; }
    public bool BackEnabled { get; set; }
    public bool NextEnabled { get; set; }

    public NavigationItem()
    {
      FinalStep = false;
    }
  }
}
