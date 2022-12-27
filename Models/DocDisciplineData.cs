using Newtonsoft.Json;

namespace DocumentLoadSanityCheckerDownload.Models
{
    public class DocDisciplineData
    {
        public string? Name { get; set; }
        [JsonProperty("SDADisciplineDocumentClass_12")]
        public List<DocClassCode>? DisciplineDocumentClass { get; set; }
    }
}
