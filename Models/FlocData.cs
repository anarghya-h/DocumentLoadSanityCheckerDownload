using Newtonsoft.Json;

namespace DocumentLoadSanityCheckerDownload.Models
{
    public class FlocData
    {
        public string? Name { get; set; }
        [JsonProperty("Floc_Level_2")]
        public string? FlocLevel2 { get; set; }
    }
}
