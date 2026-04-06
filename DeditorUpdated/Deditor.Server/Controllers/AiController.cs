using Microsoft.AspNetCore.Mvc;
using Deditor.Core.Models;
using Deditor.Core.Services.Interfaces;

namespace Deditor.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AiController : ControllerBase
    {
        private readonly IAiService _aiService;

        public AiController(IAiService aiService)
        {
            _aiService = aiService;
        }

        [HttpPost("chat")]
        public async Task<IActionResult> Chat([FromBody] ChatRequest request)
        {
            if (string.IsNullOrWhiteSpace(request?.Prompt))
                return BadRequest("No prompt provided.");

            try
            {
                var result = await _aiService.ChatAsync(request.Prompt);
                return Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                return StatusCode(500, ex.Message);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"AI request failed: {ex.Message}");
            }
        }

        [HttpPost("essentialcaps")]
        public async Task<IActionResult> EssentialCaps([FromBody] EssentialCapsRequest request)
        {
            if (string.IsNullOrWhiteSpace(request?.Text))
                return BadRequest("No text provided.");

            try
            {
                var result = await _aiService.EssentialCapsAsync(request.Text);
                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Essential Caps failed: {ex.Message}");
            }
        }
    }
}
