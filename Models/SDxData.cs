namespace DocumentLoadSanityCheckerDownload.Models
{
    public class SDxData
    {
        public RevisionCodeData? revisionCodeData { get; set; }
        public List<NameData>? originatorCompanyData { get; set; }
        public List<NameData>? docStatusCodeData { get; set; }
        public List<NameData>? languageData { get; set; }
        public List<DocDisciplineData>? disciplineData { get; set; }
        public List<NameData>? exportControlClassData { get; set; }
        public List<NameData>? mediaData { get; set; }
        public List<NameData>? securityCodes { get; set; }
        public List<LoginUserData>? Users { get; set; }
        public List<AreaCodeData>? areaCodes { get; set; }
        public List<FlocData>? flocData { get; set;}
    }
}
