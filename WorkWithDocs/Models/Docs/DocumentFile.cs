using System.ComponentModel.DataAnnotations;

namespace WorkWithDocs.Models.Docs
{
    public class DocumentFile
    {
        [Required(ErrorMessage = "Please select a file")]
        [Display(Name = "Document File")]
        public IFormFile? File { get; set; }
    }
}
