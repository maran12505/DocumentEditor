using Microsoft.AspNetCore.Mvc;
using Deditor.Core.Models;
using Deditor.Core.Services.Interfaces;

namespace Deditor.Server.Controllers
{
    [ApiController]
    [Route("api/manuscript")]
    public class ManuscriptProcessorController : ControllerBase
    {
        private readonly IManuscriptService _manuscriptService;

        public ManuscriptProcessorController(IManuscriptService manuscriptService)
        {
            _manuscriptService = manuscriptService;
        }

        [HttpPost("process")]
        public async Task<IActionResult> Process([FromBody] ManuscriptRequest request)
        {
            if (string.IsNullOrEmpty(request?.Sfdt))
                return BadRequest("No document content provided.");

            try
            {
                var result = await _manuscriptService.ProcessAsync(request.Sfdt);
                return Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                return BadRequest(ex.Message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[Manuscript] ERROR: {ex}");
                return StatusCode(500, $"Manuscript processing error: {ex.Message}");
            }
        }
    }
}
