using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Media3D;

namespace NoaaTides
{
  /// <summary>
  /// https://api.tidesandcurrents.noaa.gov/api/prod/
  /// </summary>
  internal class NoaaTidesTimeSeries
  {
    /*

     * 
     * 
     * -- p = preliminary
     * -- v = verified
     */
    //https://api.tidesandcurrents.noaa.gov/api/prod/datagetter?begin_date=20230801&end_date=20230810&station=8772471&product=water_level&datum=NAVD&time_zone=lst&units=english&application=hec.usace.army.mil&format=json

    const string baseUrl = "https://api.tidesandcurrents.noaa.gov/api/prod/datagetter";
    /// <summary>
    /// 
    /// </summary>
    /// <param name="station"></param>
    /// <param name="product">one of: water_level, hourly_height,air_temperature,water_temperature, wind, one_minute_water_level </param>
    public static async Task<DataTable> ReadTimeSeries(string station, string product, 
        DateTime startDate,DateTime endDate)
    {
      DateTime t1 = startDate;
      DataTable table = new DataTable();
      int maxDays = LookupMaxDays(product);

      while (t1 <= endDate)
      {
        DateTime t2 = t1.AddDays(30);
        if( t2 > endDate)
          t2 = endDate;

        string url = "https://api.tidesandcurrents.noaa.gov/api/prod/datagetter?"
          + "begin_date=" + t1.ToString("yyyyMMdd")
          + "&end_date=" + t2.ToString("yyyyMMdd")
          + "&station=" + station
          + "&product=" + product
          + "&datum=NAVD&time_zone=lst&units=english&application=hec.usace.army.mil&format=csv";

        string content = await Web.GetPage(url); 
        CsvFile csv = CsvFile.FromString(content);
        if (table.Rows.Count == 0)
        {
          table = csv;
        }
        else {  // append rows.
          table.Merge(csv);
        }
        t1 = t2.AddDays(1);
      }

      table.TableName = station;
      return table;
    }

    /// <summary>
    ///   1-minute interval data	Data length is limited to 4 days
    ///   6-minute interval data Data length is limited to 1 month
    ///   Hourly interval data Data length is limited to 1 year
    ///   High / Low data  Data length is limited to 1 year
    ///   Daily Means data Data length is limited to 10 years
    ///   Monthly Means data Data length is limited to 200 years
    /// </summary>
    /// <param name="product"></param>
    /// <returns></returns>
    private static int LookupMaxDays(string product)
    {//water_level, hourly_height,air_temperature,water_temperature, wind, one_minute_water_level
      switch(product)
      {
        case "one_minute_water_level":
          return 4;
        case "water_level":
          return 30;
        case "hourly_height":
          return 365;
      }
      return 365;
    }
  }
}
