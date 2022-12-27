using Newtonsoft.Json;

namespace DocumentLoadSanityCheckerDownload.Models
{
    public class OdataQueryResponse<T>
    {
        [JsonProperty(PropertyName = "@odata.count")]
        public int Count { get; set; }
        [JsonProperty(PropertyName = "value")]
        public List<T>? Value { get; set; }
        [JsonProperty(PropertyName = "@odata.nextLink")]
        public string? NextLink { get; set; }
    }
}
