using Deditor.Core.Models;
using Deditor.Core.Services.Interfaces;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Syncfusion.EJ2.DocumentEditor;
using System.Text;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Deditor.Core.Services
{
    /// <summary>
    /// Local implementation of manuscript processing — Syncfusion + OpenXML + Claude API.
    /// Used directly by MAUI; used by Server controllers to handle API requests.
    /// </summary>
    public class ManuscriptService : IManuscriptService
    {
        private readonly IConfiguration _config;
        private readonly IHttpClientFactory _httpClientFactory;

        public ManuscriptService(IConfiguration config, IHttpClientFactory httpClientFactory)
        {
            _config = config;
            _httpClientFactory = httpClientFactory;
        }

        public async Task<ManuscriptResponse> ProcessAsync(string sfdt)
        {
            if (string.IsNullOrEmpty(sfdt))
                throw new ArgumentException("No document content provided.");

            var sw = System.Diagnostics.Stopwatch.StartNew();

            // Step A: Convert SFDT -> DOCX bytes, extract text
            Console.WriteLine("[Manuscript] Converting SFDT -> DOCX...");
            Stream docxStream = WordDocument.Save(sfdt, FormatType.Docx);
            var docxMs = new MemoryStream();
            docxStream.CopyTo(docxMs);
            docxStream.Dispose();
            byte[] originalDocxBytes = docxMs.ToArray();
            docxMs.Dispose();

            var (documentText, paraTexts) = ExtractTextWithIndex(originalDocxBytes);

            if (string.IsNullOrWhiteSpace(documentText) || documentText.Length < 20)
                throw new InvalidOperationException("Document has insufficient text to analyze.");

            Console.WriteLine($"[Manuscript] Extracted {documentText.Length} chars from {paraTexts.Count} paragraphs");

            for (int dbg = 0; dbg < Math.Min(5, paraTexts.Count); dbg++)
                Console.WriteLine($"[Manuscript]   P{dbg}: \"{Truncate(paraTexts[dbg], 80)}\"");

            // Step B: Call Claude API (or mock fallback)
            var apiKey = _config["Claude:ApiKey"];
            List<ChangeItem> changes;

            if (string.IsNullOrWhiteSpace(apiKey))
            {
                Console.WriteLine("[Manuscript] No Claude API key -- using mock fallback");
                changes = GenerateMockChanges(documentText);
            }
            else
            {
                Console.WriteLine("[Manuscript] Calling Claude API...");
                try
                {
                    changes = await CallClaudeApi(apiKey, documentText);
                }
                catch (Exception apiEx)
                {
                    Console.WriteLine($"[Manuscript] Claude API failed: {apiEx.Message} -- falling back to mock");
                    changes = GenerateMockChanges(documentText);
                }
            }

            // Filter out identity changes
            changes = (changes ?? new List<ChangeItem>())
                .Where(c =>
                    !string.IsNullOrEmpty(c.OriginalText) &&
                    !string.IsNullOrEmpty(c.ModifiedText) &&
                    c.OriginalText != c.ModifiedText)
                .ToList();

            if (changes.Count == 0)
                throw new InvalidOperationException("AI did not return any actionable suggestions.");

            Console.WriteLine($"[Manuscript] Got {changes.Count} actionable changes");

            // Step C: Apply tracked changes via OpenXML
            Console.WriteLine("[Manuscript] Applying tracked changes via OpenXML...");
            var (trackedDocBytes, appliedCount) = ApplyTrackedChanges(originalDocxBytes, changes);

            Console.WriteLine($"[Manuscript] Applied {appliedCount}/{changes.Count} changes");

            if (appliedCount == 0)
                throw new InvalidOperationException("None of the AI suggestions matched text in the document.");

            // Step D: Convert tracked DOCX -> SFDT
            Console.WriteLine("[Manuscript] Converting tracked DOCX -> SFDT...");
            using var trackedStream = new MemoryStream(trackedDocBytes);
            var ej2Doc = WordDocument.Load(trackedStream, FormatType.Docx);
            string resultSfdt = JsonConvert.SerializeObject(ej2Doc);
            ej2Doc.Dispose();

            sw.Stop();
            Console.WriteLine($"[Manuscript] Done in {sw.ElapsedMilliseconds}ms, SFDT={resultSfdt.Length / 1024}KB");

            return new ManuscriptResponse
            {
                Sfdt = resultSfdt,
                ChangesApplied = appliedCount,
                TotalSuggestions = changes.Count
            };
        }

        // ══════════════════════════════════════════════════════════════════════
        // Extract text from DOCX
        // ══════════════════════════════════════════════════════════════════════
        private static (string fullText, List<string> paraTexts) ExtractTextWithIndex(byte[] docxBytes)
        {
            using var stream = new MemoryStream(docxBytes);
            using var doc = WordprocessingDocument.Open(stream, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body == null) return ("", new List<string>());

            var sb = new StringBuilder();
            var paraTexts = new List<string>();

            foreach (var para in body.Descendants<Paragraph>())
            {
                var text = GetParagraphText(para);
                paraTexts.Add(text);
                if (!string.IsNullOrEmpty(text))
                {
                    sb.Append(text);
                    sb.Append(' ');
                }
            }
            return (sb.ToString().Trim(), paraTexts);
        }

        private static string GetParagraphText(Paragraph para)
        {
            return string.Concat(para.Descendants<Text>().Select(t => t.Text));
        }

        // ══════════════════════════════════════════════════════════════════════
        // Apply tracked changes using OpenXML (w:del + w:ins)
        // ══════════════════════════════════════════════════════════════════════
        private static (byte[] docBytes, int appliedCount) ApplyTrackedChanges(
            byte[] originalDocBytes, List<ChangeItem> changes)
        {
            using var stream = new MemoryStream();
            stream.Write(originalDocBytes, 0, originalDocBytes.Length);
            stream.Position = 0;

            int appliedCount = 0;

            using (var doc = WordprocessingDocument.Open(stream, true))
            {
                var body = doc.MainDocumentPart!.Document.Body!;
                int revisionId = 100;
                string author = "AI Editor";

                foreach (var change in changes)
                {
                    if (string.IsNullOrEmpty(change.OriginalText) || string.IsNullOrEmpty(change.ModifiedText))
                        continue;

                    var paragraphs = body.Descendants<Paragraph>().ToList();
                    bool found = false;

                    foreach (var para in paragraphs)
                    {
                        var textNodes = para.Descendants<Text>().ToList();
                        if (textNodes.Count == 0) continue;

                        var fullText = new StringBuilder();
                        var nodeMap = new List<(Text node, int startPos, int endPos)>();

                        foreach (var tn in textNodes)
                        {
                            int start = fullText.Length;
                            fullText.Append(tn.Text);
                            nodeMap.Add((tn, start, fullText.Length));
                        }

                        string paraText = fullText.ToString();
                        int matchIdx = paraText.IndexOf(change.OriginalText, StringComparison.Ordinal);
                        if (matchIdx < 0) continue;

                        int matchEnd = matchIdx + change.OriginalText.Length;

                        RunProperties? matchRunProps = null;
                        foreach (var (node, startPos, endPos) in nodeMap)
                        {
                            if (endPos > matchIdx)
                            {
                                var parentRun = node.Ancestors<Run>().FirstOrDefault();
                                if (parentRun?.RunProperties != null)
                                    matchRunProps = parentRun.RunProperties.CloneNode(true) as RunProperties;
                                break;
                            }
                        }

                        string beforeText = paraText.Substring(0, matchIdx);
                        string afterText = paraText.Substring(matchEnd);

                        var pPr = para.GetFirstChild<ParagraphProperties>();
                        para.RemoveAllChildren();
                        if (pPr != null) para.AppendChild(pPr);

                        if (!string.IsNullOrEmpty(beforeText))
                        {
                            para.AppendChild(MakeRun(beforeText, matchRunProps));
                        }

                        // Deleted run
                        {
                            var delRun = new Run();
                            var delProps = matchRunProps?.CloneNode(true) as RunProperties ?? new RunProperties();
                            delProps.AppendChild(new Deleted()
                            {
                                Id = (revisionId++).ToString(),
                                Author = author,
                                Date = DateTime.UtcNow
                            });
                            delRun.RunProperties = delProps;
                            delRun.AppendChild(new DeletedText(change.OriginalText)
                                { Space = SpaceProcessingModeValues.Preserve });
                            para.AppendChild(delRun);
                        }

                        // Inserted run
                        {
                            var insRun = new Run();
                            var insProps = matchRunProps?.CloneNode(true) as RunProperties ?? new RunProperties();
                            insProps.AppendChild(new Inserted()
                            {
                                Id = (revisionId++).ToString(),
                                Author = author,
                                Date = DateTime.UtcNow
                            });
                            insRun.RunProperties = insProps;
                            insRun.AppendChild(new Text(change.ModifiedText)
                                { Space = SpaceProcessingModeValues.Preserve });
                            para.AppendChild(insRun);
                        }

                        if (!string.IsNullOrEmpty(afterText))
                        {
                            para.AppendChild(MakeRun(afterText, matchRunProps));
                        }

                        found = true;
                        appliedCount++;
                        Console.WriteLine($"[Manuscript]   TRACKED: \"{Truncate(change.OriginalText, 50)}\"");
                        break;
                    }

                    if (!found)
                    {
                        Console.WriteLine($"[Manuscript]   NOT FOUND: \"{Truncate(change.OriginalText, 60)}\"");
                    }
                }

                doc.MainDocumentPart.Document.Save();
            }

            return (stream.ToArray(), appliedCount);
        }

        private static Run MakeRun(string text, RunProperties? props)
        {
            var run = new Run();
            if (props != null) run.RunProperties = props.CloneNode(true) as RunProperties;
            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            return run;
        }

        // ══════════════════════════════════════════════════════════════════════
        // Claude API call
        // ══════════════════════════════════════════════════════════════════════
        private async Task<List<ChangeItem>> CallClaudeApi(string apiKey, string text)
        {
            if (text.Length > 100_000)
                text = text.Substring(0, 100_000);

            var prompt = "You are an expert book editor. Analyze this manuscript and find 8-12 style issues: "
                + "passive voice, tense inconsistency, long sentences (40+ words), unclear pronouns, and repetitive phrasing. "
                + "For EACH issue return EXACT original text and your suggested fix. "
                + "Return ONLY this JSON, nothing else: "
                + "{ \"changes\": [{ \"originalText\": \"exact original text from manuscript\", "
                + "\"modifiedText\": \"your corrected version\", \"reason\": \"brief explanation\" }] }";

            var model = _config["Claude:Model"] ?? "claude-sonnet-4-20250514";
            var requestBody = new
            {
                model,
                max_tokens = 4096,
                messages = new[]
                {
                    new { role = "user", content = prompt + "\n\nMANUSCRIPT:\n" + text }
                }
            };

            var client = _httpClientFactory.CreateClient();
            client.Timeout = TimeSpan.FromMinutes(2);

            var httpRequest = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
            httpRequest.Headers.Add("x-api-key", apiKey);
            httpRequest.Headers.Add("anthropic-version", "2023-06-01");
            httpRequest.Content = new StringContent(
                System.Text.Json.JsonSerializer.Serialize(requestBody),
                Encoding.UTF8,
                "application/json");

            var response = await client.SendAsync(httpRequest);
            var responseBody = await response.Content.ReadAsStringAsync();

            if (!response.IsSuccessStatusCode)
            {
                Console.WriteLine($"[Manuscript] Claude API error {response.StatusCode}: {responseBody}");
                throw new Exception($"Claude API returned {response.StatusCode}");
            }

            using var jsonDoc = System.Text.Json.JsonDocument.Parse(responseBody);
            var content = jsonDoc.RootElement.GetProperty("content");
            var textBlock = content.EnumerateArray().First(e => e.GetProperty("type").GetString() == "text");
            var aiText = textBlock.GetProperty("text").GetString() ?? "";

            var jsonText = aiText;
            var jsonMatch = System.Text.RegularExpressions.Regex.Match(aiText, @"\{[\s\S]*""changes""[\s\S]*\}");
            if (jsonMatch.Success)
                jsonText = jsonMatch.Value;

            var result = System.Text.Json.JsonSerializer.Deserialize<ChangeResponse>(jsonText, new System.Text.Json.JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });

            return result?.Changes ?? new List<ChangeItem>();
        }

        // ══════════════════════════════════════════════════════════════════════
        // Mock fallback
        // ══════════════════════════════════════════════════════════════════════
        private List<ChangeItem> GenerateMockChanges(string documentText)
        {
            var changes = new List<ChangeItem>();

            var patterns = new (string search, string replacement, string reason)[]
            {
                ("is situated", "lies",
                    "Active voice -- 'lies' is more direct than 'is situated'"),
                ("it became", "it grew into",
                    "Stronger verb -- 'grew into' is more vivid than 'became'"),
                ("is known", "has earned renown",
                    "Active voice -- removes passive 'is known' construction"),
                ("was historically", "historically served as",
                    "Active voice -- removes passive 'was' construction"),
                ("were used", "served",
                    "Active voice -- removes passive construction"),
                ("are produced", "artisans produce",
                    "Active voice -- name the agent"),
                ("was established", "emerged",
                    "Active voice -- 'emerged' is more dynamic"),
                ("was considered", "earned recognition as",
                    "Active voice -- attribute the action"),
                ("is characterized", "features",
                    "Active voice -- 'features' is more direct"),
                ("were woven", "weavers crafted",
                    "Active voice -- name the agent"),
            };

            foreach (var (search, replacement, reason) in patterns)
            {
                if (changes.Count >= 8) break;

                int idx = documentText.IndexOf(search, StringComparison.Ordinal);
                if (idx < 0) continue;

                string phrase = ExtractPhrase(documentText, idx, search.Length);
                if (string.IsNullOrEmpty(phrase)) continue;

                string modified = phrase.Replace(search, replacement);
                if (modified == phrase) continue;

                if (!documentText.Contains(phrase)) continue;

                changes.Add(new ChangeItem
                {
                    OriginalText = phrase,
                    ModifiedText = modified,
                    Reason = reason
                });

                Console.WriteLine($"[Manuscript] Mock: found \"{Truncate(phrase, 60)}\"");
            }

            Console.WriteLine($"[Manuscript] Mock generated {changes.Count} changes");
            return changes;
        }

        private static string ExtractPhrase(string text, int matchStart, int matchLength)
        {
            int start = Math.Max(0, matchStart - 80);
            int end = Math.Min(text.Length, matchStart + matchLength + 80);

            while (start > 0 && text[start] != ' ' && text[start] != '.') start--;
            if (text[start] == '.' || text[start] == ' ') start++;

            while (end < text.Length && text[end] != ' ' && text[end] != '.') end++;

            var phrase = text.Substring(start, end - start).Trim();

            if (phrase.Length < 15 || phrase.Length > 300) return "";
            return phrase;
        }

        private static string Truncate(string s, int maxLen) =>
            s.Length <= maxLen ? s : s.Substring(0, maxLen) + "...";
    }
}
