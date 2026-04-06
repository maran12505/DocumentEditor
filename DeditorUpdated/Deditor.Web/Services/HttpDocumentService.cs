using System.Net.Http.Json;
using Deditor.Core.Services.Interfaces;

namespace Deditor.Web.Services
{
    /// <summary>
    /// Web (WASM) implementation: calls the server API via HttpClient.
    /// </summary>
    public class HttpDocumentService : IDocumentService
    {
        private readonly HttpClient _http;

        public HttpDocumentService(HttpClient http)
        {
            _http = http;
        }

        public async Task<string> ImportAsync(Stream fileStream, string fileName, long fileLength)
        {
            using var formContent = new MultipartFormDataContent();
            formContent.Add(new StreamContent(fileStream), "files", fileName);

            var response = await _http.PostAsync("api/documenteditor/import", formContent);

            if (!response.IsSuccessStatusCode)
            {
                var error = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"Import failed: {error}");
            }

            return await response.Content.ReadAsStringAsync();
        }

        public Stream Save(string sfdtContent, string fileName)
        {
            // For WASM, we need to call the server synchronously-ish.
            // Use a blocking call since this is called from sync context in Razor.
            var response = _http.PostAsJsonAsync("api/documenteditor/save",
                new { Content = sfdtContent, FileName = fileName }).Result;

            if (!response.IsSuccessStatusCode)
                throw new HttpRequestException("Save failed on server.");

            var bytes = response.Content.ReadAsByteArrayAsync().Result;
            return new MemoryStream(bytes);
        }

        public string SystemClipboard(string content, string type)
        {
            var response = _http.PostAsJsonAsync("api/documenteditor/systemclipboard",
                new { Content = content, Type = type }).Result;

            if (!response.IsSuccessStatusCode)
                return string.Empty;

            return response.Content.ReadAsStringAsync().Result;
        }

        public string ServiceBase(string? imageData, string? action)
        {
            if (!string.IsNullOrEmpty(imageData))
                return imageData;
            return string.Empty;
        }
    }
}
