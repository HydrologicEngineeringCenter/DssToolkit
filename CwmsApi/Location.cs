using System.Text;
using System.Text.Json;

namespace CwmsData.Api
{
  public class Location
  {
    public string Name { get; set; } = "";
    public double Latitude { get; set; } = 0;
    public double Longitude { get; set; } = 0;
    public string TimezoneName { get; set; } = "";
    public string LocationType { get; set; } = "";
    public string LocationKind { get; set; } = "";
    public string Nation { get; set; } = "US";
    public string HorizontalDatum { get; set; } = "";
    public string OfficeId { get; set; } = "";

    private static string CreateJsonLocation(string officeId, string name, double latitude, double longitude, 
      string timezoneName, string locationKind, string nation, string horizontalDatum)
    {
      return $@"
    {{
      ""office-id"": ""{officeId}"",
      ""name"": ""{name}"",
      ""latitude"": {latitude},
      ""longitude"": {longitude},
      ""timezone-name"": ""{timezoneName}"",
      ""location-kind"": ""{locationKind}"",
      ""nation"": ""{nation}"",
      ""horizontal-datum"": ""{horizontalDatum}""
    }}";
    }

    public string ToJson()
    {
      return CreateJsonLocation(OfficeId, Name, Latitude, Longitude, TimezoneName, LocationKind, Nation, HorizontalDatum);
      //var options = new JsonSerializerOptions
      //{
      //  WriteIndented = true,
      //  PropertyNamingPolicy = new SnakeCaseNamingPolicy()
      //};
      //return JsonSerializer.Serialize(this, options);
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
