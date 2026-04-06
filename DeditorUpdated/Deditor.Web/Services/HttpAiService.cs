using System.Net.Http.Json;
using Deditor.Core.Services.Interfaces;

namespace Deditor.Web.Services
{
    /// <summary>
    /// Web (WASM) implementation: calls the server AI API via HttpClient.
    /// </summary>
    public class HttpAiService : IAiService
    {
        private readonly HttpClient _http;

        public HttpAiService(HttpClient http)
        {
            _http = http;
        }

        public async Task<string> ChatAsync(string prompt)
        {
            var response = await _http.PostAsJsonAsync("api/ai/chat", new { Prompt = prompt });

            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"AI error: {error}");
            }

            var text = await response.Content.ReadAsStringAsync();
            // Server returns the text directly (may be JSON-encoded string)
            return text.Trim('"').Replace("\\n", "\n").Replace("\\r", "");
        }

        public async Task<string> EssentialCapsAsync(string text)
        {
            var response = await _http.PostAsJsonAsync("api/ai/essentialcaps", new { Text = text });

            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"Essential Caps error: {error}");
            }

            var result = await response.Content.ReadAsStringAsync();
            return result.Trim('"').Replace("\\n", "\n").Replace("\\r", "");
        }
    }
}
