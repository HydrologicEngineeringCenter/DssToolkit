using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace CwmsApi
{
  public class Location
  {
    public string Name { get; set; } = "";
    public double Latitude { get; set; } = 0;
    public double Longitude { get; set; } = 0;
    public bool Active { get; set; } = true;
    public string PublicName { get; set; } = "";
    public string LongName { get; set; } = "";
    public string Description { get; set; } = "";
    public string TimezoneName { get; set; } = "";
    public string LocationType { get; set; } = "";
    public string LocationKind { get; set; } = "";
    public string Nation { get; set; } = "US";
    public string StateInitial { get; set; } = "";
    public string CountyName { get; set; } = "";
    public string NearestCity { get; set; } = "";
    public string HorizontalDatum { get; set; } = "";
    public double PublishedLongitude { get; set; } = 0;
    public double PublishedLatitude { get; set; } = 0;
    public string VerticalDatum { get; set; } = "";
    public double Elevation { get; set; } = 0;
    public string MapLabel { get; set; } = "";
    public string BoundingOfficeId { get; set; } = "";
    public string OfficeId { get; set; } = "";
    public string ToJson()
    {
      var options = new JsonSerializerOptions
      {
        WriteIndented = true,
        PropertyNamingPolicy = new SnakeCaseNamingPolicy()
      };
      return JsonSerializer.Serialize(this, options);
    }
  }
  public class SnakeCaseNamingPolicy : JsonNamingPolicy
  {
    public override string ConvertName(string name)
    {
      StringBuilder result = new StringBuilder();
      for (int i = 0; i < name.Length; i++)
      {
        if (char.IsUpper(name[i]) && i > 0)
        {
          result.Append('-');
        }
        result.Append(char.ToLower(name[i]));
      }
      return result.ToString();
    }
  }
}
