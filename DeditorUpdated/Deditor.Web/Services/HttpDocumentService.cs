using System.Net;
using System.Net.Http.Json;
using Deditor.Core.Models;
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

        public async Task<Stream> SaveAsync(string sfdtContent, string fileName)
        {
            var response = await _http.PostAsJsonAsync("api/documenteditor/save",
                new { Content = sfdtContent, FileName = fileName });

            if (!response.IsSuccessStatusCode)
                throw new HttpRequestException("Save failed on server.");

            var bytes = await response.Content.ReadAsByteArrayAsync();
            return new MemoryStream(bytes);
        }

        public async Task<string> SystemClipboardAsync(string content, string type)
        {
            var response = await _http.PostAsJsonAsync("api/documenteditor/systemclipboard",
                new { Content = content, Type = type });

            if (!response.IsSuccessStatusCode)
                return string.Empty;

            return await response.Content.ReadAsStringAsync();
        }

        public string ServiceBase(string? imageData, string? action)
        {
            if (!string.IsNullOrEmpty(imageData))
                return imageData;
            return string.Empty;
        }

        public async Task<byte[]> UpdateIPubCustomXmlAsync(byte[] docxBytes, IPubMetaDto dto)
        {
            Console.WriteLine("[HttpDocumentService.UpdateIPubCustomXml] ===== START =====");
            try
            {
                using var formContent = new MultipartFormDataContent();
                formContent.Add(new ByteArrayContent(docxBytes), "files", "document.docx");
                var json = System.Text.Json.JsonSerializer.Serialize(dto);
                formContent.Add(new StringContent(json), "dto");

                Console.WriteLine($"[HttpDocumentService.UpdateIPubCustomXml] POST update-customxml, bytes={docxBytes.Length}, dto.length={json.Length}");
                var response = await _http.PostAsync("api/documenteditor/update-customxml", formContent);
                Console.WriteLine($"[HttpDocumentService.UpdateIPubCustomXml] response status = {(int)response.StatusCode}");
                if (!response.IsSuccessStatusCode)
                {
                    Console.WriteLine("[HttpDocumentService.UpdateIPubCustomXml] non-success — returning original bytes.");
                    return docxBytes;
                }
                var bytes = await response.Content.ReadAsByteArrayAsync();
                Console.WriteLine($"[HttpDocumentService.UpdateIPubCustomXml] received {bytes.Length} bytes.");
                Console.WriteLine("[HttpDocumentService.UpdateIPubCustomXml] ===== END =====");
                return bytes;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[HttpDocumentService.UpdateIPubCustomXml] EXCEPTION: {ex.Message}");
                return docxBytes;
            }
        }

        public async Task<IPubMetaDto?> ExtractIPubMetaAsync(Stream docxStream)
        {
            Console.WriteLine("[HttpDocumentService.ExtractIPubMeta] ===== START =====");
            try
            {
                using var formContent = new MultipartFormDataContent();
                formContent.Add(new StreamContent(docxStream), "files", "document.docx");

                Console.WriteLine("[HttpDocumentService.ExtractIPubMeta] POST api/documenteditor/extract-metadata");
                var response = await _http.PostAsync("api/documenteditor/extract-metadata", formContent);
                Console.WriteLine($"[HttpDocumentService.ExtractIPubMeta] response status = {(int)response.StatusCode} {response.StatusCode}");

                if (response.StatusCode == HttpStatusCode.NoContent)
                {
                    Console.WriteLine("[HttpDocumentService.ExtractIPubMeta] 204 NoContent — no metadata extracted.");
                    Console.WriteLine("[HttpDocumentService.ExtractIPubMeta] ===== END =====");
                    return null;
                }
                if (!response.IsSuccessStatusCode)
                {
                    var err = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"[HttpDocumentService.ExtractIPubMeta] non-success body: {err}");
                    Console.WriteLine("[HttpDocumentService.ExtractIPubMeta] ===== END =====");
                    return null;
                }

                var dto = await response.Content.ReadFromJsonAsync<IPubMetaDto>();
                Console.WriteLine($"[HttpDocumentService.ExtractIPubMeta] DTO received. Publisher='{dto?.Publisher}' BookTitle='{dto?.BookTitle}' Isbn='{dto?.Isbn}'");
                Console.WriteLine("[HttpDocumentService.ExtractIPubMeta] ===== END =====");
                return dto;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[HttpDocumentService.ExtractIPubMeta] EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                Console.WriteLine("[HttpDocumentService.ExtractIPubMeta] ===== END (exception) =====");
                return null;
            }
        }
    }
}
