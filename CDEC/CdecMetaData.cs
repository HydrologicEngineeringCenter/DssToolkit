using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tools;

namespace CDEC
{
  /// <summary>
  /// 
  /// </summary>
  internal class CdecMetaData
  {

     static async Task<DataTable> GetStations()
    {
      string content = await Web.GetPage("https://cdec.water.ca.gov/reportapp/javareports?name=RealStations");


      return new DataTable(content);
    }
    ///
    ///


  }
}
