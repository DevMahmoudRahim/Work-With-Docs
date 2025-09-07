using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using WorkWithDocs.DTO.Docs;
using WorkWithDocs.Models;
using WorkWithDocs.Models.Docs;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Presentation;
//using DocumentFormat.OpenXml.Drawing;

namespace WorkWithDocs.Controllers
{
    public class DocumentController : Controller
    {
        private readonly ILogger<DocumentController> logger;
        private readonly IWebHostEnvironment env;

        public DocumentController(ILogger<DocumentController> logger, IWebHostEnvironment env)
        {
            this.logger = logger;
            this.env = env;
        }

        public async Task<IActionResult> Index()
        {
            return View(new DocumentFile());
        }

        [HttpPost]
        public async Task<IActionResult> UploadDocumant(DocumentFile model) 
        {
            if(!ModelState.IsValid || model.File.FileName is null)
            {
                return View(nameof(Index));
            }

            try
            {
                var allowedExtensions = new[] { ".doc", ".docx", ".pptx", ".ppt" };
                var fileExtension = Path.GetExtension(model.File.FileName).ToLower();

                if(!allowedExtensions.Contains(fileExtension))
                {
                    ModelState.AddModelError(string.Empty, "Invalid file type. Only .doc, .docx, .pptx, .ppt files are allowed.");
                    return View(nameof(Index));
                }

                var uploadPath = Path.Combine(env.WebRootPath, "uploads", model.File.FileName);

                if(!Directory.Exists(uploadPath))
                    Directory.CreateDirectory(uploadPath);

                var fileName = model.File.FileName;
                var filePath = Path.Combine(uploadPath, fileName);

                using(var stream = new FileStream(filePath, FileMode.Create))
                {
                    await model.File.CopyToAsync(stream);
                }

                string content = "";
                if (fileExtension == ".docx" || fileExtension == ".doc")
                    content = ExtractWordContent(filePath);

                else if (fileExtension == ".pptx" || fileExtension == ".ppt")
                    content = ExtractPowerPointContent(filePath);

                var documentFile = new DocumentModel
                {
                    Content = content,
                    FileName = fileName,
                    FileType = fileExtension,
                    FilePath = filePath,
                    IsSuccess = true,
                    Message = "Document processed successfully"
                };
                return View("DocumentEditor", documentFile);
            }

            catch(Exception ex)
            {
                logger.LogError(ex, "Error updating document");
                ModelState.AddModelError(string.Empty, "An error occurred while updating the document.");
                return View(nameof(Index));
            }
        }

        [HttpGet]
        public async Task<IActionResult> DownloadDocument(string fileName)
        {
            try
            {
                var uploadsPath = Path.Combine(env.WebRootPath, "uploads");
                var filePath = Path.Combine(uploadsPath, fileName);

                if (!System.IO.File.Exists(filePath))
                {
                    return NotFound();
                }

                var fileBytes = System.IO.File.ReadAllBytes(filePath);
                var originalFileName = fileName.Substring(fileName.IndexOf('_') + 1);

                return File(fileBytes, GetContentType(fileName), originalFileName);
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error downloading document");
                return BadRequest("Error downloading file.");
            }
        }

        [HttpPost]
        public async Task<IActionResult> UpdateDocumentContent([FromBody] DocumentUpdateDto model)
        {
            try
            {
                var uploadPath = Path.Combine(env.WebRootPath, "uploads");
                var fullPath = Path.Combine(env.WebRootPath, model.FilePath);

                if (!System.IO.File.Exists(fullPath))
                    return Json(new { success = false, message = "File not found." });

                var fileExtension = Path.GetExtension(model.FilePath).ToLower();

                if (fileExtension == ".docx" || fileExtension == ".doc")
                {
                    UpdateWordContent(fullPath, model.Content);
                }
                else if (fileExtension == ".pptx" || fileExtension == ".ppt")
                {
                    UpdatePowerPointContent(fullPath, model.Content);
                }

                return Json(new { success = true, message = "Document updated successfully!" });
            }

            catch (Exception ex)
            {
                logger.LogError(ex, "Error updating document");
                return Json(new { success = false, message = "An error occurred while updating the document." });
            }
        }

        private string ExtractWordContent(string filePath)
        {
            try
            {
                using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
                {
                    return wordDoc.MainDocumentPart.Document.Body.InnerText;
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error extracting Word content from {FilePath}", filePath);
                return $"Error reading Word document content: {ex.Message}";
            }
        }

        private string ExtractPowerPointContent(string filePath)
        {
            var content = new System.Text.StringBuilder();

            try
            {
                using (PresentationDocument pptDoc = PresentationDocument.Open(filePath, false))
                {
                    PresentationPart presentationPart = pptDoc.PresentationPart;
                    if (presentationPart == null || presentationPart.Presentation == null)
                    {
                        return "Presentation is empty.";
                    }

                    foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
                    {
                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);

                        var allText = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Text>().Select(t => t.Text);
                        content.AppendLine(string.Join(Environment.NewLine, allText));
                    }
                }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error extracting PowerPoint content from {FilePath}", filePath);
                return $"Error reading PowerPoint presentation content: {ex.Message}";
            }

            return content.ToString();
        }

        private void UpdatePowerPointContent(string filePath, string content)
        {
            var textFilePath = filePath.Replace(Path.GetExtension(filePath), ".txt");
            System.IO.File.WriteAllText(textFilePath, content);
        }

        private void UpdateWordContent(string filePath, string content)
        {
            try
            {
                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite))
                {
                    using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(fileStream, true))
                    {
                        MainDocumentPart mainPart = wordDoc.MainDocumentPart;
                        if (mainPart == null || mainPart.Document.Body == null)
                        {
                            throw new InvalidOperationException("Word document body not found.");
                        }

                        Body body = mainPart.Document.Body;

                        Paragraph newParagraph = new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(content)));
                        body.AppendChild(newParagraph);
                    }
            }
            catch (Exception ex)
            {
                logger.LogError(ex, "Error updating Word document content for {FilePath}", filePath);
                throw; 
            }
        }

        private string GetContentType(string fileName)
        {
            var extension = Path.GetExtension(fileName).ToLowerInvariant();
            return extension switch
            {
                ".doc" => "application/msword",
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".ppt" => "application/vnd.ms-powerpoint",
                ".pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                _ => "application/octet-stream"
            };
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

    }
}
