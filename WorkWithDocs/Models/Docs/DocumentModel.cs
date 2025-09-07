using System.ComponentModel.DataAnnotations;

namespace WorkWithDocs.Models.Docs
{
    public class DocumentModel
    {
        public string FileName { get; set; } = string.Empty;
        public string FileType { get; set; } = string.Empty;
        public string Content { get; set; } = string.Empty;
        public string FilePath { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public bool IsSuccess { get; set; }

    }
}
