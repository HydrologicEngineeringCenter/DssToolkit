using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DssExcel
{
  public abstract class ValidationVM:BaseVM
  {
    public abstract bool Validate(out string errorMessage);
  }
}
