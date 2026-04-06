using Deditor.Core.Services.Interfaces;
using Microsoft.Extensions.Configuration;
using System.Text;
using System.Text.Json;

namespace Deditor.Core.Services
{
    /// <summary>
    /// Local implementation of AI operations — calls OpenRouter API directly.
    /// Used directly by MAUI; used by Server controllers to handle API requests.
    /// </summary>
    public class AiService : IAiService
    {
        private readonly IConfiguration _config;
        private readonly IHttpClientFactory _httpFactory;

        public AiService(IConfiguration config, IHttpClientFactory httpFactory)
        {
            _config = config;
            _httpFactory = httpFactory;
        }

        public async Task<string> ChatAsync(string prompt)
        {
            if (string.IsNullOrWhiteSpace(prompt))
                throw new ArgumentException("No prompt provided.");

            var apiKey = _config["OpenRouter:ApiKey"];
            if (string.IsNullOrEmpty(apiKey))
                throw new InvalidOperationException("OpenRouter API key not configured. Add it to appsettings.json under OpenRouter:ApiKey.");

            var http = _httpFactory.CreateClient();

            // OpenRouter — free models, OpenAI-compatible API
            http.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
            http.DefaultRequestHeaders.Add("HTTP-Referer", "http://localhost:5278");
            http.DefaultRequestHeaders.Add("X-Title", "Deditor");

            var body = new
            {
                model = "openrouter/auto",
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
                        content = prompt
                    }
                }
            };

            var json = JsonSerializer.Serialize(body);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await http.PostAsync("https://openrouter.ai/api/v1/chat/completions", content);

            if (!response.IsSuccessStatusCode)
            {
                var err = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"OpenRouter API error ({response.StatusCode}): {err}");
            }

            var raw = await response.Content.ReadAsStringAsync();
            using var doc = JsonDocument.Parse(raw);

            var text = doc.RootElement
                .GetProperty("choices")[0]
                .GetProperty("message")
                .GetProperty("content")
                .GetString();

            return text ?? string.Empty;
        }

        public async Task<string> EssentialCapsAsync(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                throw new ArgumentException("No text provided.");

            var prompt = $"""
                Apply "Essential Caps" to the text below.
                Rules: Capitalize ONLY important content words (nouns, technical/scientific terms,
                proper nouns, domain keywords). Keep lowercase: articles, prepositions, conjunctions,
                auxiliary verbs. Return ONLY the transformed text, nothing else.

                Text: {text}
                """;

            return await ChatAsync(prompt);
        }
    }
}
