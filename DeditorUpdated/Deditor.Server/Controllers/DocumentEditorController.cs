using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http.Timeouts;
using Deditor.Core.Models;
using Deditor.Core.Services;
using Deditor.Core.Services.Interfaces;

namespace Deditor.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentEditorController : ControllerBase
    {
        private readonly IDocumentService _documentService;

        public DocumentEditorController(IDocumentService documentService)
        {
            _documentService = documentService;
        }

        [DisableRequestSizeLimit]
        [RequestFormLimits(MultipartBodyLengthLimit = long.MaxValue)]
        [HttpPost("import")]
        public async Task<IActionResult> Import(IFormFile files)
        {
            if (files == null || files.Length == 0)
                return BadRequest("No file received.");

            try
            {
                using var ms = new MemoryStream();
                await files.CopyToAsync(ms);
                ms.Position = 0;

                var sfdt = await _documentService.ImportAsync(ms, files.FileName, files.Length);
                return Ok(sfdt);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Import] ERROR: {ex}");
                return StatusCode(500, $"Import error: {ex.Message}");
            }
        }

        [HttpPost("save")]
        public IActionResult Save([FromBody] SaveParameter data)
        {
            if (data == null || string.IsNullOrEmpty(data.Content))
                return BadRequest("No content received.");
            try
            {
                var docStream = _documentService.Save(data.Content, data.FileName ?? "document.docx");
                var mime = DocumentService.GetMimeType(data.FileName ?? "document.docx");
                return File(docStream, mime, data.FileName ?? "document.docx");
            }
            catch (Exception ex) { return StatusCode(500, $"Save error: {ex.Message}"); }
        }

        [HttpPost("systemclipboard")]
        public IActionResult SystemClipboard([FromBody] CustomParameter param)
        {
            if (string.IsNullOrEmpty(param?.Content)) return Ok(string.Empty);
            var sfdt = _documentService.SystemClipboard(param.Content, param.Type ?? "html");
            return Ok(sfdt);
        }

        [HttpPost("restricteditor")]
        public IActionResult RestrictEditor([FromBody] RestrictParameter param) => Ok(false);

        [HttpPost("loaddefaultdictionary")]
        public IActionResult LoadDefaultDictionary() => Ok(string.Empty);

        [HttpPost("spellcheck")]
        public IActionResult SpellCheck([FromBody] SpellCheckParameter param)
            => Ok(new { HasSpellingError = false, Suggestions = Array.Empty<string>() });

        [HttpPost("spellcheckbypage")]
        public IActionResult SpellCheckByPage([FromBody] object param) => Ok(string.Empty);

        [HttpPost("")]
        [Consumes("application/json")]
        public IActionResult ServiceBase([FromBody] MetafileParam? param)
        {
            if (param == null)
                return Ok(string.Empty);

            var result = _documentService.ServiceBase(param.imageData, param.action);
            return Ok(result);
        }
    }
}
