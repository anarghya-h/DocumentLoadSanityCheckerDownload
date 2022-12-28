using System.ComponentModel.DataAnnotations;

namespace DocumentLoadSanityCheckerDownload.Models
{
    public class PlantCodeData
    {
        public string ? Name { get; set; }
        [Required(ErrorMessage ="Please select a plant")]
        public string? UID { get; set; }
    }
}
