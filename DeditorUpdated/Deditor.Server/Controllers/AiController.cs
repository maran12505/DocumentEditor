using Microsoft.AspNetCore.Mvc;
using System.Text;
using System.Text.Json;

namespace Deditor.Server.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AiController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly IHttpClientFactory _httpFactory;

        public AiController(IConfiguration config, IHttpClientFactory httpFactory)
        {
            _config = config;
            _httpFactory = httpFactory;
        }

        // POST api/ai/chat
        // Proofreading, copyediting, and free-form Q&A via Groq (free tier)
        [HttpPost("chat")]
        public async Task<IActionResult> Chat([FromBody] ChatRequest request)
        {
            if (string.IsNullOrWhiteSpace(request?.Prompt))
                return BadRequest("No prompt provided.");

            var apiKey = _config["OpenRouter:ApiKey"];
            if (string.IsNullOrEmpty(apiKey))
                return StatusCode(500, "OpenRouter API key not configured. Add it to appsettings.json under OpenRouter:ApiKey.");

            try
            {
                var http = _httpFactory.CreateClient();

                // OpenRouter — free models, OpenAI-compatible API
                http.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
                http.DefaultRequestHeaders.Add("HTTP-Referer", "http://localhost:5278");
                http.DefaultRequestHeaders.Add("X-Title", "Deditor");

                var body = new
                {
                    model = "openrouter/auto",  // free model on OpenRouter
                    max_tokens = 1024,
                    temperature = 0.3,
                    messages = new[]
                    {
                        new
                        {
                            role = "system",
                            content = """
                                You are an expert copyeditor and proofreader specializing in
                                academic, scientific, and technical publishing. Be concise and
                                practical. When proofreading, list each issue clearly with the
                                corrected version. When copyediting, show the improved version
                                first (labeled "Improved version:") then briefly explain changes.
                                Use plain text — avoid excessive markdown.
                                """
                        },
                        new
                        {
                            role = "user",
                            content = request.Prompt
                        }
                    }
                };

                var json = JsonSerializer.Serialize(body);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // OpenRouter endpoint — OpenAI-compatible
                var response = await http.PostAsync("https://openrouter.ai/api/v1/chat/completions", content);

                if (!response.IsSuccessStatusCode)
                {
                    var err = await response.Content.ReadAsStringAsync();
                    return StatusCode((int)response.StatusCode, $"Groq API error: {err}");
                }

                var raw = await response.Content.ReadAsStringAsync();
                using var doc = JsonDocument.Parse(raw);

                // OpenAI-compatible response: choices[0].message.content
                var text = doc.RootElement
                    .GetProperty("choices")[0]
                    .GetProperty("message")
                    .GetProperty("content")
                    .GetString();

                return Ok(text);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"AI request failed: {ex.Message}");
            }
        }

        // POST api/ai/essentialcaps
        [HttpPost("essentialcaps")]
        public async Task<IActionResult> EssentialCaps([FromBody] EssentialCapsRequest request)
        {
            if (string.IsNullOrWhiteSpace(request?.Text))
                return BadRequest("No text provided.");

            var prompt = $"""
                Apply "Essential Caps" to the text below.
                Rules: Capitalize ONLY important content words (nouns, technical/scientific terms,
                proper nouns, domain keywords). Keep lowercase: articles, prepositions, conjunctions,
                auxiliary verbs. Return ONLY the transformed text, nothing else.

                Text: {request.Text}
                """;

            return await Chat(new ChatRequest { Prompt = prompt });
        }
    }

    public class ChatRequest         { public string? Prompt { get; set; } }
    public class EssentialCapsRequest { public string? Text  { get; set; } }
}