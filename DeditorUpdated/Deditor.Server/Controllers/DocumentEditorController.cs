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

        [DisableRequestSizeLimit]
        [RequestFormLimits(MultipartBodyLengthLimit = long.MaxValue)]
        [HttpPost("extract-metadata")]
        public async Task<IActionResult> ExtractMetadata(IFormFile files)
        {
            Console.WriteLine("[ExtractMetadata] ===== START endpoint =====");
            if (files == null || files.Length == 0)
            {
                Console.WriteLine("[ExtractMetadata] No file received. 400.");
                Console.WriteLine("[ExtractMetadata] ===== END =====");
                return BadRequest("No file received.");
            }
            Console.WriteLine($"[ExtractMetadata] file='{files.FileName}' size={files.Length} bytes");
            try
            {
                using var ms = new MemoryStream();
                await files.CopyToAsync(ms);
                ms.Position = 0;
                Console.WriteLine($"[ExtractMetadata] Buffered {ms.Length} bytes. Calling ExtractIPubMetaAsync...");
                var dto = await _documentService.ExtractIPubMetaAsync(ms);
                if (dto == null)
                {
                    Console.WriteLine("[ExtractMetadata] DTO is null — returning 204 NoContent.");
                    Console.WriteLine("[ExtractMetadata] ===== END =====");
                    return NoContent();
                }
                Console.WriteLine("[ExtractMetadata] DTO returned — sending 200 OK with JSON.");
                Console.WriteLine("[ExtractMetadata] ===== END =====");
                return Ok(dto);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ExtractMetadata] ERROR: {ex}");
                Console.WriteLine("[ExtractMetadata] ===== END (exception) =====");
                return StatusCode(500, $"Extract error: {ex.Message}");
            }
        }

        [DisableRequestSizeLimit]
        [RequestFormLimits(MultipartBodyLengthLimit = long.MaxValue)]
        [HttpPost("update-customxml")]
        public async Task<IActionResult> UpdateCustomXml([FromForm] IFormFile files, [FromForm] string dto)
        {
            Console.WriteLine("[UpdateCustomXml] ===== START endpoint =====");
            if (files == null || files.Length == 0)
            {
                Console.WriteLine("[UpdateCustomXml] No file received. 400.");
                return BadRequest("No file received.");
            }
            if (string.IsNullOrWhiteSpace(dto))
            {
                Console.WriteLine("[UpdateCustomXml] No DTO JSON. 400.");
                return BadRequest("No metadata DTO received.");
            }
            Console.WriteLine($"[UpdateCustomXml] file='{files.FileName}' size={files.Length}, dto length={dto.Length}");

            IPubMetaDto? parsed;
            try
            {
                parsed = System.Text.Json.JsonSerializer.Deserialize<IPubMetaDto>(dto,
                    new System.Text.Json.JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            }
            catch (Exception jsonEx)
            {
                Console.WriteLine($"[UpdateCustomXml] DTO parse failed: {jsonEx.Message}");
                return BadRequest("Invalid DTO JSON.");
            }
            if (parsed == null)
            {
                Console.WriteLine("[UpdateCustomXml] Parsed DTO is null.");
                return BadRequest("Invalid DTO.");
            }

            try
            {
                using var ms = new MemoryStream();
                await files.CopyToAsync(ms);
                var inBytes = ms.ToArray();
                Console.WriteLine($"[UpdateCustomXml] Buffered {inBytes.Length} bytes. Calling UpdateIPubCustomXmlAsync...");

                var outBytes = await _documentService.UpdateIPubCustomXmlAsync(inBytes, parsed);
                Console.WriteLine($"[UpdateCustomXml] Service returned {outBytes.Length} bytes. Sending file response.");
                Console.WriteLine("[UpdateCustomXml] ===== END =====");
                return File(outBytes,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    files.FileName);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[UpdateCustomXml] ERROR: {ex}");
                Console.WriteLine("[UpdateCustomXml] ===== END (exception) =====");
                return StatusCode(500, $"UpdateCustomXml error: {ex.Message}");
            }
        }

        [HttpPost("save")]
        public async Task<IActionResult> Save([FromBody] SaveParameter data)
        {
            if (data == null || string.IsNullOrEmpty(data.Content))
                return BadRequest("No content received.");
            try
            {
                var docStream = await _documentService.SaveAsync(data.Content, data.FileName ?? "document.docx");
                var mime = DocumentService.GetMimeType(data.FileName ?? "document.docx");
                return File(docStream, mime, data.FileName ?? "document.docx");
            }
            catch (Exception ex) { return StatusCode(500, $"Save error: {ex.Message}"); }
        }

        [HttpPost("systemclipboard")]
        public async Task<IActionResult> SystemClipboard([FromBody] CustomParameter param)
        {
            if (string.IsNullOrEmpty(param?.Content)) return Ok(string.Empty);
            var sfdt = await _documentService.SystemClipboardAsync(param.Content, param.Type ?? "html");
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
