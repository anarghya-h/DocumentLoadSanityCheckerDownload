using Newtonsoft.Json;

namespace DocumentLoadSanityCheckerDownload.Models
{
    public class LoginUserData
    {
        public string? Name { get; set; }
        [JsonProperty(PropertyName = "Address_email")]
        public string? Email { get; set; }
    }
}
