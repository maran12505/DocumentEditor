using Microsoft.AspNetCore.Mvc;
using Syncfusion.EJ2.DocumentEditor;
using Newtonsoft.Json;

namespace Deditor.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentEditorController : ControllerBase
    {
        // POST api/documenteditor/import
        // Upload DOCX/DOC/RTF/TXT → returns SFDT JSON string
        [HttpPost("import")]
        public IActionResult Import(IFormFile files)
        {
            if (files == null || files.Length == 0)
                return BadRequest("No file received.");

            try
            {
                using var ms = new MemoryStream();
                files.CopyTo(ms);
                ms.Position = 0;

                // Load stream → WordDocument SFDT DOM
                WordDocument document = WordDocument.Load(ms, GetFormatType(files.FileName));

                // Serialize the SFDT DOM to JSON for the client editor
                string sfdt = JsonConvert.SerializeObject(document);
                document.Dispose();

                return Ok(sfdt);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Import error: {ex.Message}");
            }
        }

        // POST api/documenteditor/save
        // SFDT JSON string → DOCX file download
        [HttpPost("save")]
        public IActionResult Save([FromBody] SaveParameter data)
        {
            if (data == null || string.IsNullOrEmpty(data.Content))
                return BadRequest("No content received.");

            try
            {
                // IMPORTANT: Must use the TWO-argument overload so the return type is Stream.
                // The single-argument overload returns Syncfusion.DocIO.DLS.WordDocument (not a Stream).
                FormatType fmt = GetFormatType(data.FileName ?? "document.docx");
                Stream docStream = WordDocument.Save(data.Content, fmt);

                var contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
                return File(docStream, contentType, data.FileName ?? "document.docx");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Save error: {ex.Message}");
            }
        }

        // POST api/documenteditor/systemclipboard
        // Paste-with-formatting: HTML/RTF clipboard content → SFDT
        [HttpPost("systemclipboard")]
        public IActionResult SystemClipboard([FromBody] CustomParameter param)
        {
            if (string.IsNullOrEmpty(param?.Content))
                return Ok(string.Empty);

            try
            {
                WordDocument document = WordDocument.LoadString(param.Content, GetClipboardFormat(param.Type));
                string sfdt = JsonConvert.SerializeObject(document);
                document.Dispose();
                return Ok(sfdt);
            }
            catch
            {
                return Ok(string.Empty);
            }
        }

        // POST api/documenteditor/restricteditor
        [HttpPost("restricteditor")]
        public IActionResult RestrictEditor([FromBody] RestrictParameter param)
        {
            return Ok(false);
        }

        // ── Helpers ──────────────────────────────────────────────────────────

        private static FormatType GetFormatType(string? fileName) =>
            Path.GetExtension(fileName ?? "").ToLowerInvariant() switch
            {
                ".docx" => FormatType.Docx,
                ".doc"  => FormatType.Doc,
                ".rtf"  => FormatType.Rtf,
                ".txt"  => FormatType.Txt,
                _       => FormatType.Docx
            };

        private static FormatType GetClipboardFormat(string? type) =>
            type?.ToLowerInvariant() switch
            {
                "html" => FormatType.Html,
                "rtf"  => FormatType.Rtf,
                _      => FormatType.Html
            };
    }

    // ── Parameter models ─────────────────────────────────────────────────────

    public class SaveParameter
    {
        public string? Content  { get; set; }
        public string? FileName { get; set; }
    }

    public class CustomParameter
    {
        public string? Content { get; set; }
        public string? Type    { get; set; }
    }

    public class RestrictParameter
    {
        public string? PasswordBase64 { get; set; }
        public string? SaltBase64     { get; set; }
        public int     SpinCount      { get; set; }
    }
}