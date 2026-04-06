using System.Net.Http.Json;
using Deditor.Core.Models;
using Deditor.Core.Services.Interfaces;

namespace Deditor.Web.Services
{
    /// <summary>
    /// Web (WASM) implementation: calls the server manuscript API via HttpClient.
    /// </summary>
    public class HttpManuscriptService : IManuscriptService
    {
        private readonly HttpClient _http;

        public HttpManuscriptService(HttpClient http)
        {
            _http = http;
        }

        public async Task<ManuscriptResponse> ProcessAsync(string sfdt)
        {
            var response = await _http.PostAsJsonAsync("api/manuscript/process", new { Sfdt = sfdt });

            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"Manuscript processing failed: {error}");
            }

            var result = await response.Content.ReadFromJsonAsync<ManuscriptResponse>();
            return result ?? throw new InvalidOperationException("No result returned from server.");
        }
    }
}
