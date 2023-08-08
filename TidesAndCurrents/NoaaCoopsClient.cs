using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TidesAndCurrents
{
  internal class NoaaCoopsClient
  {
    public DataTable StationList()
    {
      string url = @"https://api.tidesandcurrents.noaa.gov/dpapi/prod/webapi/product/.xml?name=toptenwaterlevels&units=english";
      var rval = new DataTable();

      rval.ReadXml()
    }
  }
}
