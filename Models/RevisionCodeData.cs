namespace DocumentLoadSanityCheckerDownload.Models
{
    public class RevisionCodeData
    {
        public string? Name { get; set; }
        public string? Major_Seq { get; set; }
        public string? Minor_Seq { get; set; }
        public List<string>? MajorRevision { get; set; }
        public List<string>? MinorRevision { get; set; }
    }
}
