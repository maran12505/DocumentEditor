using Deditor.Core.Models;
using Deditor.Core.Services.Interfaces;
using DocumentFormat.OpenXml.Packaging;
using Syncfusion.EJ2.DocumentEditor;
using Newtonsoft.Json;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace Deditor.Core.Services
{
    /// <summary>
    /// Local implementation of document operations using Syncfusion + SkiaSharp.
    /// Used directly by MAUI; used by Server controllers to handle API requests.
    /// </summary>
    public class DocumentService : IDocumentService
    {
        private static readonly Dictionary<string, string> _cache = new();

        // Max dimension for converted images — keeps SFDT size reasonable
        private const int MaxImageDimension = 1200;
        private const long JpegQuality = 70L;
        private const long MaxFileSizeBytes = 100 * 1024 * 1024; // 100 MB

        public async Task<string> ImportAsync(Stream fileStream, string fileName, long fileLength)
        {
            if (fileLength > MaxFileSizeBytes)
                throw new InvalidOperationException(
                    $"File too large ({fileLength / 1024 / 1024}MB). Maximum supported size is {MaxFileSizeBytes / 1024 / 1024}MB.");

            var sw = System.Diagnostics.Stopwatch.StartNew();
            var key = fileName + fileLength;

            if (_cache.TryGetValue(key, out var cached))
            {
                Console.WriteLine($"[Import] Cache hit: {fileName}");
                return cached;
            }

            using var ms = new MemoryStream();
            await fileStream.CopyToAsync(ms);
            ms.Position = 0;

            var ext = Path.GetExtension(fileName).ToLowerInvariant();
            Console.WriteLine($"[Import] {fileName} ({fileLength} bytes, ext={ext})");

            // Pre-process: convert unsupported images
            Stream streamForSyncfusion = ms;
            FormatType loadFormat = GetFormatType(fileName);

            if (ext == ".doc" || ext == ".rtf")
            {
                Console.WriteLine($"[Import] Converting {ext} -> .docx for image processing...");
                var tempDoc = WordDocument.Load(ms, loadFormat);
                string tempSfdt = JsonConvert.SerializeObject(tempDoc);
                tempDoc.Dispose();
                // Convert SFDT -> DOCX bytes
                Stream docxRaw = WordDocument.Save(tempSfdt, FormatType.Docx);
                var docxStream = new MemoryStream();
                docxRaw.CopyTo(docxStream);
                docxRaw.Dispose();
                docxStream.Position = 0;

                var converted = ConvertUnsupportedImages(docxStream);
                if (converted != null)
                {
                    streamForSyncfusion = converted;
                    docxStream.Dispose();
                    loadFormat = FormatType.Docx;
                    Console.WriteLine($"[Import] {ext} images converted, loading as DOCX");
                }
                else
                {
                    docxStream.Position = 0;
                    streamForSyncfusion = docxStream;
                    loadFormat = FormatType.Docx;
                    Console.WriteLine($"[Import] {ext} -> DOCX (no image conversion needed)");
                }
            }
            else if (ext == ".docx")
            {
                var converted = ConvertUnsupportedImages(ms);
                if (converted != null)
                {
                    streamForSyncfusion = converted;
                    Console.WriteLine($"[Import] Pre-processed: {ms.Length / 1024}KB -> {converted.Length / 1024}KB");
                }
                else
                {
                    ms.Position = 0;
                }
            }

            // Load with Syncfusion
            WordDocument document;

            if (ext == ".xml")
            {
                using var reader = new StreamReader(streamForSyncfusion);
                var xmlContent = reader.ReadToEnd();
                document = WordDocument.LoadString(xmlContent, FormatType.Html);
            }
            else
            {
                document = WordDocument.Load(streamForSyncfusion, loadFormat);
            }

            string sfdt = JsonConvert.SerializeObject(document);
            document.Dispose();

            // Dispose the converted stream if we created one
            if (streamForSyncfusion != ms)
                streamForSyncfusion.Dispose();

            // Quick diagnostic
            var imgCount = Regex.Matches(sfdt, @"""imageString""\s*:\s*""[^""]{10,}").Count;
            Console.WriteLine($"[Import] SFDT: {sfdt.Length / 1024}KB, images found in SFDT: {imgCount}");

            _cache[key] = sfdt;
            sw.Stop();
            Console.WriteLine($"[Import] Done in {sw.ElapsedMilliseconds} ms");

            return sfdt;
        }

        public Stream Save(string sfdtContent, string fileName)
        {
            FormatType fmt = GetFormatType(fileName);
            return WordDocument.Save(sfdtContent, fmt);
        }

        public string SystemClipboard(string content, string type)
        {
            if (string.IsNullOrEmpty(content)) return string.Empty;
            try
            {
                WordDocument document = WordDocument.LoadString(content, GetClipboardFormat(type));
                string sfdt = JsonConvert.SerializeObject(document);
                document.Dispose();
                return sfdt;
            }
            catch { return string.Empty; }
        }

        public string ServiceBase(string? imageData, string? action)
        {
            if (!string.IsNullOrEmpty(imageData))
            {
                Console.WriteLine($"[ServiceBase] imageData: {imageData.Length} chars, action={action}");
                return imageData;
            }
            return string.Empty;
        }

        // ══════════════════════════════════════════════════════════════════════
        // CustomXml -> iPubEdit metadata extractor.
        // Best-effort: any failure returns null. Never alters caller's stream.
        // ══════════════════════════════════════════════════════════════════════
        public async Task<IPubMetaDto?> ExtractIPubMetaAsync(Stream docxStream)
        {
            Console.WriteLine("[ExtractIPubMeta] ===== START =====");
            if (docxStream == null || !docxStream.CanRead)
            {
                Console.WriteLine("[ExtractIPubMeta] Stream is null or not readable. ABORT.");
                Console.WriteLine("[ExtractIPubMeta] ===== END (null) =====");
                return null;
            }
            Console.WriteLine($"[ExtractIPubMeta] Stream OK. CanSeek={docxStream.CanSeek}");

            using var ms = new MemoryStream();
            try
            {
                if (docxStream.CanSeek) docxStream.Position = 0;
                await docxStream.CopyToAsync(ms);
                ms.Position = 0;
                Console.WriteLine($"[ExtractIPubMeta] Buffered {ms.Length} bytes into memory.");
            }
            catch (Exception copyEx)
            {
                Console.WriteLine($"[ExtractIPubMeta] Buffer copy failed: {copyEx.Message}");
                Console.WriteLine("[ExtractIPubMeta] ===== END (copy error) =====");
                return null;
            }

            try
            {
                using var doc = WordprocessingDocument.Open(ms, false);
                Console.WriteLine("[ExtractIPubMeta] Opened DOCX with WordprocessingDocument.");
                var main = doc.MainDocumentPart;
                if (main == null)
                {
                    Console.WriteLine("[ExtractIPubMeta] MainDocumentPart is null. ABORT.");
                    Console.WriteLine("[ExtractIPubMeta] ===== END (no main part) =====");
                    return null;
                }

                var partList = main.CustomXmlParts.ToList();
                Console.WriteLine($"[ExtractIPubMeta] customXml part count = {partList.Count}");
                if (partList.Count == 0)
                {
                    Console.WriteLine("[ExtractIPubMeta] No customXml parts present in DOCX. Returning null.");
                    Console.WriteLine("[ExtractIPubMeta] ===== END (no customXml) =====");
                    return null;
                }

                var dto = new IPubMetaDto();
                bool any = false;
                int partIndex = 0;
                int matchedFields = 0;

                foreach (var part in partList)
                {
                    partIndex++;
                    Console.WriteLine($"[ExtractIPubMeta] -- Part #{partIndex} uri={part.Uri} contentType={part.ContentType}");

                    XDocument xdoc;
                    string rawPreview = "";
                    try
                    {
                        using var s = part.GetStream(FileMode.Open, FileAccess.Read);
                        using var reader = new StreamReader(s);
                        var raw = reader.ReadToEnd();
                        rawPreview = raw.Length > 500 ? raw.Substring(0, 500) + "..." : raw;
                        Console.WriteLine($"[ExtractIPubMeta]    raw length = {raw.Length}");
                        Console.WriteLine($"[ExtractIPubMeta]    raw preview: {rawPreview}");
                        xdoc = XDocument.Parse(raw);
                    }
                    catch (Exception partEx)
                    {
                        Console.WriteLine($"[ExtractIPubMeta]    Failed to parse part: {partEx.Message}. Skipping.");
                        continue;
                    }

                    Console.WriteLine($"[ExtractIPubMeta]    root element = {xdoc.Root?.Name.LocalName} (ns={xdoc.Root?.Name.NamespaceName})");

                    int descendantCount = 0;
                    int partMatches = 0;
                    foreach (var el in xdoc.Descendants())
                    {
                        descendantCount++;
                        var localName = el.Name.LocalName;
                        if (string.IsNullOrEmpty(localName)) continue;

                        // Special-case <article ...> — JATS root carries metadata in attributes:
                        //   article-type → ArticleType, dtd-version → Dtd, xml:lang → Language
                        if (localName.Equals("article", StringComparison.OrdinalIgnoreCase))
                        {
                            var articleType = (string?)el.Attribute("article-type");
                            if (!string.IsNullOrWhiteSpace(articleType) && string.IsNullOrEmpty(dto.ArticleType))
                            {
                                dto.ArticleType = articleType.Trim(); any = true; partMatches++; matchedFields++;
                                Console.WriteLine($"[ExtractIPubMeta]    [MATCH attr]     <article article-type='{articleType}'> -> ArticleType");
                            }
                            var dtdVersion = (string?)el.Attribute("dtd-version");
                            if (!string.IsNullOrWhiteSpace(dtdVersion) && string.IsNullOrEmpty(dto.Dtd))
                            {
                                dto.Dtd = dtdVersion.Trim(); any = true; partMatches++; matchedFields++;
                                Console.WriteLine($"[ExtractIPubMeta]    [MATCH attr]     <article dtd-version='{dtdVersion}'> -> Dtd");
                            }
                            var lang = (string?)el.Attribute(XNamespace.Xml + "lang");
                            if (!string.IsNullOrWhiteSpace(lang) && string.IsNullOrEmpty(dto.Language))
                            {
                                dto.Language = lang.Trim(); any = true; partMatches++; matchedFields++;
                                Console.WriteLine($"[ExtractIPubMeta]    [MATCH attr]     <article xml:lang='{lang}'> -> Language");
                            }
                            // fall through — no element-text mapping for <article>
                        }

                        // Special-case <journal-id journal-id-type="publisher|publisher-id|...">
                        if (localName.Equals("journal-id", StringComparison.OrdinalIgnoreCase) && !el.HasElements)
                        {
                            var v = (el.Value ?? string.Empty).Trim();
                            if (v.Length > 0 && !IsTokenPlaceholder(v))
                            {
                                var idType = ((string?)el.Attribute("journal-id-type") ?? "").ToLowerInvariant();
                                // "publisher", "publisher-id", "doi", or unspecified — treat as JournalName short code
                                if (string.IsNullOrEmpty(dto.JournalName))
                                {
                                    dto.JournalName = v; any = true; partMatches++; matchedFields++;
                                    Console.WriteLine($"[ExtractIPubMeta]    [MATCH]          <journal-id journal-id-type='{idType}'> = '{v}' -> JournalName");
                                }
                            }
                            continue;
                        }

                        // Special-case <pub-id pub-id-type="doi|publisher-id|art-access-id|manuscript|...">
                        if (localName.Equals("pub-id", StringComparison.OrdinalIgnoreCase) && !el.HasElements)
                        {
                            var v = (el.Value ?? string.Empty).Trim();
                            if (v.Length == 0 || IsTokenPlaceholder(v))
                            {
                                Console.WriteLine($"[ExtractIPubMeta]    [skip placeholder/empty] <pub-id> = '{v}'");
                                continue;
                            }
                            var idType = ((string?)el.Attribute("pub-id-type") ?? "").ToLowerInvariant();
                            if (idType == "doi")
                            {
                                if (string.IsNullOrEmpty(dto.Doi)) { dto.Doi = v; any = true; partMatches++; matchedFields++; }
                                Console.WriteLine($"[ExtractIPubMeta]    [MATCH]          <pub-id pub-id-type='doi'> = '{v}' -> Doi");
                            }
                            else if (idType == "publisher-id" || idType == "art-access-id" || idType == "manuscript" || idType == "pii" || idType == "other")
                            {
                                if (string.IsNullOrEmpty(dto.ArticleId)) { dto.ArticleId = v; any = true; partMatches++; matchedFields++; }
                                Console.WriteLine($"[ExtractIPubMeta]    [MATCH]          <pub-id pub-id-type='{idType}'> = '{v}' -> ArticleId");
                            }
                            else
                            {
                                Console.WriteLine($"[ExtractIPubMeta]    [no map]         <pub-id pub-id-type='{idType}'> = '{v}'");
                            }
                            continue;
                        }

                        // Special-case <issn pub-type="ppub|epub">
                        if (localName.Equals("issn", StringComparison.OrdinalIgnoreCase) && !el.HasElements)
                        {
                            var v = (el.Value ?? string.Empty).Trim();
                            if (v.Length == 0 || IsTokenPlaceholder(v))
                            {
                                Console.WriteLine($"[ExtractIPubMeta]    [skip placeholder/empty] <issn> = '{v}'");
                                continue;
                            }
                            var pubType = (string?)el.Attribute("pub-type") ?? string.Empty;
                            if (pubType.Equals("epub", StringComparison.OrdinalIgnoreCase))
                            {
                                if (string.IsNullOrEmpty(dto.EIssn)) { dto.EIssn = v; any = true; partMatches++; matchedFields++; }
                                Console.WriteLine($"[ExtractIPubMeta]    [MATCH]          <issn pub-type=epub> = '{v}' -> EIssn");
                            }
                            else // ppub or unspecified
                            {
                                if (string.IsNullOrEmpty(dto.PIssn)) { dto.PIssn = v; any = true; partMatches++; matchedFields++; }
                                Console.WriteLine($"[ExtractIPubMeta]    [MATCH]          <issn pub-type={pubType}> = '{v}' -> PIssn");
                            }
                            continue;
                        }

                        if (IPubFieldMap.TryGetValue(localName, out var setter))
                        {
                            if (el.HasElements)
                            {
                                Console.WriteLine($"[ExtractIPubMeta]    [skip container] <{localName}> has child elements");
                                continue;
                            }
                            var value = (el.Value ?? string.Empty).Trim();
                            if (value.Length == 0)
                            {
                                Console.WriteLine($"[ExtractIPubMeta]    [skip empty]     <{localName}>");
                                continue;
                            }
                            if (IsTokenPlaceholder(value))
                            {
                                Console.WriteLine($"[ExtractIPubMeta]    [skip placeholder] <{localName}> = '{value}'");
                                continue;
                            }
                            if (_dropdownLikeKeys.Contains(localName) && IsDropdownNoise(value))
                            {
                                Console.WriteLine($"[ExtractIPubMeta]    [skip dropdown-noise] <{localName}> = '{value}'");
                                continue;
                            }
                            setter(dto, value);
                            any = true;
                            partMatches++;
                            matchedFields++;
                            Console.WriteLine($"[ExtractIPubMeta]    [MATCH]          <{localName}> = '{value}'");
                        }
                        else
                        {
                            Console.WriteLine($"[ExtractIPubMeta]    [no map]         <{localName}>");
                        }
                    }
                    Console.WriteLine($"[ExtractIPubMeta]    part summary: descendants={descendantCount}, matched={partMatches}");
                }

                Console.WriteLine($"[ExtractIPubMeta] Total matched fields across all parts = {matchedFields}");
                if (any)
                {
                    Console.WriteLine($"[ExtractIPubMeta] DTO populated: Publisher='{dto.Publisher}' BookTitle='{dto.BookTitle}' Isbn='{dto.Isbn}' Doi='{dto.Doi}' JournalName='{dto.JournalName}' Year='{dto.Year}'");
                }
                else
                {
                    Console.WriteLine("[ExtractIPubMeta] No mapped fields found — returning null.");
                }
                Console.WriteLine("[ExtractIPubMeta] ===== END =====");
                return any ? dto : null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ExtractIPubMeta] Failed: {ex.GetType().Name}: {ex.Message}");
                Console.WriteLine($"[ExtractIPubMeta] StackTrace: {ex.StackTrace}");
                Console.WriteLine("[ExtractIPubMeta] ===== END (exception) =====");
                return null;
            }
        }

        // Helper: return true if value is empty or a semicolon-wrapped template
        // token like ";isbnpbk;", ";journaltitle;" — these are placeholders the
        // upstream pipeline emits when the real value is missing.
        private static bool IsTokenPlaceholder(string v) =>
            v.Length >= 2 && v.StartsWith(';') && v.EndsWith(';');

        // Helper: return true for "noise" values that should not be assigned to
        // dropdown-bound fields. "0", "00", whitespace etc. mean "no selection"
        // upstream — assigning them clobbers the model default and the UI just
        // shows "Select an option" anyway.
        private static bool IsDropdownNoise(string v) =>
            string.IsNullOrWhiteSpace(v) || v.Trim().TrimStart('0').Length == 0;

        // Element local-names whose extracted value drives a UI dropdown.
        // Filtered through IsDropdownNoise before being applied.
        private static readonly HashSet<string> _dropdownLikeKeys =
            new(StringComparer.OrdinalIgnoreCase)
            {
                "icore-pagination-platform", "PaginationPlatform",
                "itracks-jobcardid", "itracks-subjobcardid", "JobCard", "JobCardNumber",
                "country", "Country", "loc", "publisher-loc",
                "ProcessType",
            };

        // First-non-empty-wins setter helpers (keep the first concrete value seen
        // for a field; ignore subsequent occurrences). Repeated elements like
        // multiple <isbn> in the same part should not clobber a good first value.
        private static void SetIfEmpty(IPubMetaDto d, ref string? target, string v)
        {
            if (string.IsNullOrEmpty(target)) target = v;
        }

        // Maps customXml local element names (case-insensitive) to setters on
        // the DTO. Add aliases freely — extractor logic doesn't change.
        internal static readonly Dictionary<string, Action<IPubMetaDto, string>> IPubFieldMap =
            new(StringComparer.OrdinalIgnoreCase)
            {
                // ── Generic / canonical names ────────────────────────────────
                ["XmlType"]            = (d, v) => { if (string.IsNullOrEmpty(d.XmlType))            d.XmlType = v; },
                ["Publisher"]          = (d, v) => { if (string.IsNullOrEmpty(d.Publisher))          d.Publisher = v; },
                ["JobCard"]            = (d, v) => { if (string.IsNullOrEmpty(d.JobCard) || d.JobCard == "0") d.JobCard = v; },
                ["JobCardNumber"]      = (d, v) => { if (string.IsNullOrEmpty(d.JobCard) || d.JobCard == "0") d.JobCard = v; },
                ["CeTemplate"]         = (d, v) => { if (string.IsNullOrEmpty(d.CeTemplate))         d.CeTemplate = v; },
                ["BookOrJournalId"]    = (d, v) => { if (string.IsNullOrEmpty(d.BookOrJournalId))    d.BookOrJournalId = v; },
                ["BookId"]             = (d, v) => { if (string.IsNullOrEmpty(d.BookOrJournalId))    d.BookOrJournalId = v; },
                ["JournalId"]          = (d, v) => { if (string.IsNullOrEmpty(d.BookOrJournalId))    d.BookOrJournalId = v; },
                ["Copyedit"]           = (d, v) => { if (!d.Copyedit.HasValue)   d.Copyedit   = ParseBool(v); },
                ["PreEditing"]         = (d, v) => { if (!d.PreEditing.HasValue) d.PreEditing = ParseBool(v); },
                ["Testing"]            = (d, v) => { if (!d.Testing.HasValue)    d.Testing    = ParseBool(v); },
                ["Dtd"]                = (d, v) => { if (string.IsNullOrEmpty(d.Dtd))                d.Dtd = v; },
                ["DTD"]                = (d, v) => { if (string.IsNullOrEmpty(d.Dtd))                d.Dtd = v; },
                ["BookTitle"]          = (d, v) => { if (string.IsNullOrEmpty(d.BookTitle))          d.BookTitle = v; },
                ["Title"]              = (d, v) => { if (string.IsNullOrEmpty(d.BookTitle))          d.BookTitle = v; },
                ["Isbn"]               = (d, v) => { if (string.IsNullOrEmpty(d.Isbn))               d.Isbn = v; },
                ["ISBN"]               = (d, v) => { if (string.IsNullOrEmpty(d.Isbn))               d.Isbn = v; },
                ["Month"]              = (d, v) => { if (string.IsNullOrEmpty(d.Month))              d.Month = v; },
                ["Year"]               = (d, v) => { if (string.IsNullOrEmpty(d.Year))               d.Year = v; },
                ["PaginationPlatform"] = (d, v) => { if (string.IsNullOrEmpty(d.PaginationPlatform)) d.PaginationPlatform = v; },
                ["Country"]            = (d, v) => { if (string.IsNullOrEmpty(d.Country) || d.Country == "None") d.Country = v; },
                ["ArticleType"]        = (d, v) => { if (string.IsNullOrEmpty(d.ArticleType))        d.ArticleType = v; },
                ["Doi"]                = (d, v) => { if (string.IsNullOrEmpty(d.Doi))                d.Doi = v; },
                ["DOI"]                = (d, v) => { if (string.IsNullOrEmpty(d.Doi))                d.Doi = v; },
                ["Language"]           = (d, v) => { if (string.IsNullOrEmpty(d.Language))           d.Language = v; },
                ["ProcessType"]        = (d, v) => { if (string.IsNullOrEmpty(d.ProcessType))        d.ProcessType = v; },
                ["JournalName"]        = (d, v) => { if (string.IsNullOrEmpty(d.JournalName))        d.JournalName = v; },
                ["JournalTitle"]       = (d, v) => { if (string.IsNullOrEmpty(d.JournalTitle))       d.JournalTitle = v; },
                ["PIssn"]              = (d, v) => { if (string.IsNullOrEmpty(d.PIssn))              d.PIssn = v; },
                ["pISSN"]              = (d, v) => { if (string.IsNullOrEmpty(d.PIssn))              d.PIssn = v; },
                ["EIssn"]              = (d, v) => { if (string.IsNullOrEmpty(d.EIssn))              d.EIssn = v; },
                ["eISSN"]              = (d, v) => { if (string.IsNullOrEmpty(d.EIssn))              d.EIssn = v; },
                ["Copyrights"]         = (d, v) => { if (string.IsNullOrEmpty(d.Copyrights))         d.Copyrights = v; },
                ["Copyright"]          = (d, v) => { if (string.IsNullOrEmpty(d.Copyrights))         d.Copyrights = v; },
                ["PublisherImprint"]   = (d, v) => { if (string.IsNullOrEmpty(d.PublisherImprint))   d.PublisherImprint = v; },
                ["Imprint"]            = (d, v) => { if (string.IsNullOrEmpty(d.PublisherImprint))   d.PublisherImprint = v; },
                ["ArticleId"]          = (d, v) => { if (string.IsNullOrEmpty(d.ArticleId))          d.ArticleId = v; },
                ["DocumentType"]       = (d, v) => { if (string.IsNullOrEmpty(d.DocumentType))       d.DocumentType = v; },

                // ── iCoRe / iTracks / JATS-style schema (real DOCX names) ────
                ["icore-publisher"]          = (d, v) => { if (string.IsNullOrEmpty(d.Publisher))          d.Publisher = v; },
                ["icore-xmltype"]            = (d, v) => { if (string.IsNullOrEmpty(d.XmlType))            d.XmlType = v; },
                ["icore-dtd"]                = (d, v) => { if (string.IsNullOrEmpty(d.Dtd))                d.Dtd = v; },
                ["icore-isCopyEditing"]      = (d, v) => { if (!d.Copyedit.HasValue)   d.Copyedit   = ParseBool(v); },
                ["icore-isPreEditing"]       = (d, v) => { if (!d.PreEditing.HasValue) d.PreEditing = ParseBool(v); },
                ["icore-TestingEXE"]         = (d, v) => { if (!d.Testing.HasValue)    d.Testing    = ParseBool(v); },
                ["icore-CopyEditingTemplate"] = (d, v) => { if (string.IsNullOrEmpty(d.CeTemplate))        d.CeTemplate = v; },
                // Note: <icore-CEType> intentionally NOT mapped — it is a type code
                // (e.g. "TE" = Technical Edit), not a CE template name. Mapping it
                // would clobber the real <icore-CopyEditingTemplate> value.
                ["icore-pagination-platform"] = (d, v) => { if (string.IsNullOrEmpty(d.PaginationPlatform)) d.PaginationPlatform = v; },
                ["icore-doi"]                = (d, v) => { if (string.IsNullOrEmpty(d.Doi))                d.Doi = v; },
                ["icore-Category"]           = (d, v) => { if (string.IsNullOrEmpty(d.ArticleType))        d.ArticleType = v; },
                ["icore-copyrightstatement"] = (d, v) => { if (string.IsNullOrEmpty(d.Copyrights))         d.Copyrights = v; },
                ["icore-copyrightholder"]    = (d, v) => { if (string.IsNullOrEmpty(d.Copyrights))         d.Copyrights = v; },
                ["icore-copyrightyear"]      = (d, v) => { /* covered by Copyrights statement; ignore */ },

                ["itracks-jobcardid"]    = (d, v) => { if (string.IsNullOrEmpty(d.JobCard) || d.JobCard == "0") d.JobCard = v; },
                ["itracks-subjobcardid"] = (d, v) => { /* not used */ },

                ["book-id"]              = (d, v) => { if (string.IsNullOrEmpty(d.BookOrJournalId))    d.BookOrJournalId = v; },
                ["book-title"]           = (d, v) => { if (string.IsNullOrEmpty(d.BookTitle))          d.BookTitle = v; },
                ["publisher-name"]       = (d, v) => { if (string.IsNullOrEmpty(d.Publisher))          d.Publisher = v; },
                // publisher-loc is the publisher's city/region — distinct from
                // the document's country-of-publication. Map to PublisherLoc.
                ["publisher-loc"]        = (d, v) => { if (string.IsNullOrEmpty(d.PublisherLoc)) d.PublisherLoc = v; },
                ["journalcode"]          = (d, v) => { if (string.IsNullOrEmpty(d.JournalName))        d.JournalName = v; },
                ["journaltitle"]         = (d, v) => { if (string.IsNullOrEmpty(d.JournalTitle))       d.JournalTitle = v; },
                ["ChapterType"]          = (d, v) => { if (string.IsNullOrEmpty(d.DocumentType))       d.DocumentType = v; },
                ["DeliverableType"]      = (d, v) => { if (string.IsNullOrEmpty(d.DocumentType))       d.DocumentType = v; },
                ["country"]              = (d, v) => { if (string.IsNullOrEmpty(d.Country) || d.Country == "None") d.Country = v; },
                ["language"]             = (d, v) => { if (string.IsNullOrEmpty(d.Language))           d.Language = v; },
                ["isbn"]                 = (d, v) => { if (string.IsNullOrEmpty(d.Isbn))               d.Isbn = v; },
                ["year"]                 = (d, v) => { if (string.IsNullOrEmpty(d.Year))               d.Year = v; },
                ["month"]                = (d, v) => { if (string.IsNullOrEmpty(d.Month))              d.Month = v; },

                // ── JATS canonical (per iCoRe-Journal.xsd) ───────────────────
                ["journal-title"]        = (d, v) => { if (string.IsNullOrEmpty(d.JournalTitle))       d.JournalTitle = v; },
                ["abbrev-journal-title"] = (d, v) => { if (string.IsNullOrEmpty(d.JournalName))        d.JournalName = v; },
                ["journal-subtitle"]     = (d, v) => { /* not in UI */ },
                ["pub"]                  = (d, v) => { if (string.IsNullOrEmpty(d.Publisher))          d.Publisher = v; },
                ["loc"]                  = (d, v) => { if (string.IsNullOrEmpty(d.Country) || d.Country == "None") d.Country = v; },
                ["copyright-statement"]  = (d, v) => { if (string.IsNullOrEmpty(d.Copyrights))         d.Copyrights = v; },
                ["copyright-holder"]     = (d, v) => { if (string.IsNullOrEmpty(d.Copyrights))         d.Copyrights = v; },
                ["copyright-year"]       = (d, v) => { /* covered by copyright-statement */ },
                ["mth"]                  = (d, v) => { if (string.IsNullOrEmpty(d.Month))              d.Month = v; },
                ["yr"]                   = (d, v) => { if (string.IsNullOrEmpty(d.Year))               d.Year = v; },
                ["article-id"]           = (d, v) => { if (string.IsNullOrEmpty(d.ArticleId))          d.ArticleId = v; },
                ["object-id"]            = (d, v) => { if (string.IsNullOrEmpty(d.ArticleId))          d.ArticleId = v; },
                ["doi"]                  = (d, v) => { if (string.IsNullOrEmpty(d.Doi))                d.Doi = v; },
                ["int-doi"]              = (d, v) => { if (string.IsNullOrEmpty(d.Doi))                d.Doi = v; },
                ["subject"]              = (d, v) => { if (string.IsNullOrEmpty(d.ArticleType))        d.ArticleType = v; },

                // ── Newly-mapped XSD elements (Book + Journal) ────────────────
                ["article-title"]        = (d, v) => { if (string.IsNullOrEmpty(d.ArticleTitle))       d.ArticleTitle = v; },
                ["chapter-title"]        = (d, v) => { if (string.IsNullOrEmpty(d.ChapterTitle))       d.ChapterTitle = v; },
                ["volume"]               = (d, v) => { if (string.IsNullOrEmpty(d.Volume))             d.Volume = v; },
                ["issue"]                = (d, v) => { if (string.IsNullOrEmpty(d.Issue))              d.Issue = v; },
                ["fpage"]                = (d, v) => { if (string.IsNullOrEmpty(d.FirstPage))          d.FirstPage = v; },
                ["lpage"]                = (d, v) => { if (string.IsNullOrEmpty(d.LastPage))           d.LastPage = v; },
                ["elocation-id"]         = (d, v) => { if (string.IsNullOrEmpty(d.ElocationId))        d.ElocationId = v; },
                ["edition"]              = (d, v) => { if (string.IsNullOrEmpty(d.Edition))            d.Edition = v; },
                ["series"]               = (d, v) => { if (string.IsNullOrEmpty(d.Series))             d.Series = v; },
                ["day"]                  = (d, v) => { if (string.IsNullOrEmpty(d.Day))                d.Day = v; },
            };

        private static bool ParseBool(string v) =>
            v.Equals("true", StringComparison.OrdinalIgnoreCase) ||
            v == "1" ||
            v.Equals("yes", StringComparison.OrdinalIgnoreCase);

        // ══════════════════════════════════════════════════════════════════════
        // CustomXml write-back: persist DTO values into the iCore customXml part.
        // Existing elements are updated in place. Missing elements are created
        // as children of the iCore root. Multi-element fields like <isbn> update
        // the first non-placeholder; <issn> uses pub-type. Namespace is preserved.
        // Returns the modified DOCX bytes, or the original bytes on any failure.
        // ══════════════════════════════════════════════════════════════════════
        public async Task<byte[]> UpdateIPubCustomXmlAsync(byte[] docxBytes, IPubMetaDto dto)
        {
            Console.WriteLine("[UpdateMeta] ===== START =====");
            if (docxBytes == null || docxBytes.Length == 0)
            {
                Console.WriteLine("[UpdateMeta] Empty input bytes — returning unchanged.");
                Console.WriteLine("[UpdateMeta] ===== END =====");
                return docxBytes ?? Array.Empty<byte>();
            }
            if (dto == null)
            {
                Console.WriteLine("[UpdateMeta] DTO is null — returning unchanged.");
                Console.WriteLine("[UpdateMeta] ===== END =====");
                return docxBytes;
            }

            // Work on a copy; never mutate caller's array.
            var ms = new MemoryStream();
            await ms.WriteAsync(docxBytes, 0, docxBytes.Length);
            ms.Position = 0;
            Console.WriteLine($"[UpdateMeta] Buffered {ms.Length} bytes.");

            try
            {
                using (var doc = WordprocessingDocument.Open(ms, true))
                {
                    var main = doc.MainDocumentPart;
                    if (main == null)
                    {
                        Console.WriteLine("[UpdateMeta] MainDocumentPart is null. ABORT.");
                        Console.WriteLine("[UpdateMeta] ===== END =====");
                        return docxBytes;
                    }

                    // Find the iCore customXml part (root local-name == "icore").
                    // Prefer /customXml/item3.xml when multiple iCore parts exist.
                    CustomXmlPart? icorePart = null;
                    XDocument? icoreDoc = null;

                    foreach (var part in main.CustomXmlParts)
                    {
                        XDocument? candidate;
                        try
                        {
                            using var s = part.GetStream(FileMode.Open, FileAccess.Read);
                            candidate = XDocument.Load(s);
                        }
                        catch (Exception readEx)
                        {
                            Console.WriteLine($"[UpdateMeta] Failed to read part {part.Uri}: {readEx.Message}");
                            continue;
                        }
                        var rootName = candidate.Root?.Name.LocalName;
                        Console.WriteLine($"[UpdateMeta] Inspecting part uri={part.Uri} root=<{rootName}>");
                        if (string.Equals(rootName, "icore", StringComparison.OrdinalIgnoreCase))
                        {
                            // Prefer item3 when more than one iCore part is present.
                            if (icorePart == null || part.Uri.OriginalString.EndsWith("item3.xml", StringComparison.OrdinalIgnoreCase))
                            {
                                icorePart = part;
                                icoreDoc = candidate;
                            }
                        }
                    }

                    bool autoCreated = false;
                    if (icorePart == null || icoreDoc == null || icoreDoc.Root == null)
                    {
                        Console.WriteLine("[AutoCreateXML] No customXml found — creating new part");
                        try
                        {
                            // Use the string content-type overload — matches what the
                            // upstream pipeline emits ("application/xml") so the new
                            // part round-trips through extraction without surprises.
                            icorePart = main.AddCustomXmlPart("application/xml");
                            icoreDoc = BuildEmptyICoreTemplate();
                            using (var initStream = icorePart.GetStream(FileMode.Create, FileAccess.Write))
                            {
                                icoreDoc.Save(initStream);
                            }
                            Console.WriteLine($"[AutoCreateXML] item created at uri={icorePart.Uri}");
                            // Re-load from the part so we operate on the canonical XDocument.
                            using (var rs = icorePart.GetStream(FileMode.Open, FileAccess.Read))
                            {
                                icoreDoc = XDocument.Load(rs);
                            }
                            autoCreated = true;
                            Console.WriteLine("[AutoCreateXML] item1.xml created");
                        }
                        catch (Exception createEx)
                        {
                            Console.WriteLine($"[AutoCreateXML] Failed to create customXml part: {createEx.Message}");
                            Console.WriteLine("[UpdateMeta] ===== END (auto-create failed) =====");
                            return docxBytes;
                        }
                    }
                    else
                    {
                        Console.WriteLine($"[UpdateMeta] Found iCore XML part: {icorePart.Uri}");
                    }

                    var root = icoreDoc.Root;
                    var ns = root.Name.Namespace; // preserve namespace if any

                    int updated = 0;
                    int created = 0;

                    // Iterate the write-map and apply each non-null DTO field.
                    foreach (var (localName, getter) in _icoreWriteMap)
                    {
                        var value = getter(dto);
                        if (value == null) continue;

                        // Empty string is a legitimate "clear the field" intent — keep it.
                        // (The UI shows "" for empty bound values; if the user clears DOI
                        // they expect <icore-doi></icore-doi> to be empty in the file.)
                        var existing = FindFirstByLocalName(root, localName);
                        if (existing != null)
                        {
                            if (existing.Value != value)
                            {
                                Console.WriteLine($"[UpdateMeta] Updating <{localName}> '{existing.Value}' → '{value}'");
                                existing.Value = value;
                                updated++;
                            }
                            else
                            {
                                Console.WriteLine($"[UpdateMeta] Unchanged <{localName}> = '{value}' (skip)");
                            }
                        }
                        else
                        {
                            // Don't auto-create when the value is empty — empty + missing
                            // means there's nothing to record.
                            if (string.IsNullOrEmpty(value))
                            {
                                Console.WriteLine($"[UpdateMeta] Skip create empty <{localName}>");
                                continue;
                            }
                            // Create as child of root, preserving any default namespace.
                            var newEl = ns == XNamespace.None
                                ? new XElement(localName, value)
                                : new XElement(ns + localName, value);
                            root.Add(newEl);
                            Console.WriteLine($"[UpdateMeta] Created <{localName}> = '{value}'");
                            created++;
                        }
                    }

                    // Special-case ISSN: <issn pub-type="ppub"|"epub"> — update by attribute.
                    UpdateOrCreateIssn(root, ns, "ppub", dto.PIssn, ref updated, ref created);
                    UpdateOrCreateIssn(root, ns, "epub", dto.EIssn, ref updated, ref created);

                    Console.WriteLine($"[UpdateMeta] Summary: updated={updated}, created={created}");

                    // Write back into the customXml part.
                    using (var ws = icorePart.GetStream(FileMode.Create, FileAccess.Write))
                    {
                        icoreDoc.Save(ws);
                    }
                    Console.WriteLine("[UpdateMeta] XML saved successfully");
                    if (autoCreated)
                        Console.WriteLine("[AutoCreateXML] Metadata written successfully");
                } // dispose closes/saves the package

                Console.WriteLine("[UpdateMeta] DOCX repacked.");
                Console.WriteLine("[UpdateMeta] ===== END =====");
                return ms.ToArray();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[UpdateMeta] EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                Console.WriteLine($"[UpdateMeta] StackTrace: {ex.StackTrace}");
                Console.WriteLine("[UpdateMeta] ===== END (exception) =====");
                return docxBytes;
            }
            finally
            {
                ms.Dispose();
            }
        }

        // Find first descendant whose local name matches (case-insensitive),
        // ignoring namespace.
        private static XElement? FindFirstByLocalName(XElement scope, string localName) =>
            scope.DescendantsAndSelf()
                 .FirstOrDefault(e => string.Equals(e.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase));

        // Builds the empty iCore customXml template used when a DOCX has no
        // customXml at all. Mirrors the structure required by the upstream
        // pipeline so a freshly-created part round-trips through extraction.
        private static XDocument BuildEmptyICoreTemplate()
        {
            return new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XElement("icore",
                    new XElement("icore-info",
                        new XElement("icore-publisher"),
                        new XElement("icore-xmltype"),
                        new XElement("icore-dtd"),
                        new XElement("icore-isCopyEditing"),
                        new XElement("icore-isPreEditing"),
                        new XElement("icore-CEType"),
                        new XElement("icore-CopyEditingTemplate"),
                        new XElement("icore-doi"),
                        new XElement("book",
                            new XElement("book-meta",
                                new XElement("book-id"),
                                new XElement("book-title-group",
                                    new XElement("book-title")),
                                new XElement("isbn"),
                                new XElement("pub-date",
                                    new XElement("year"),
                                    new XElement("month")),
                                new XElement("language"))))));
        }

        private static void UpdateOrCreateIssn(XElement root, XNamespace ns, string pubType, string? value,
                                               ref int updated, ref int created)
        {
            if (value == null) return;
            // Find existing <issn pub-type="...">
            var existing = root.Descendants()
                .Where(e => string.Equals(e.Name.LocalName, "issn", StringComparison.OrdinalIgnoreCase))
                .FirstOrDefault(e => string.Equals((string?)e.Attribute("pub-type"), pubType, StringComparison.OrdinalIgnoreCase));

            if (existing != null)
            {
                if (existing.Value != value)
                {
                    Console.WriteLine($"[UpdateMeta] Updating <issn pub-type='{pubType}'> '{existing.Value}' → '{value}'");
                    existing.Value = value;
                    updated++;
                }
                return;
            }
            if (string.IsNullOrEmpty(value)) return;

            // Create a new <issn pub-type="..."> at root level
            var el = ns == XNamespace.None ? new XElement("issn", value) : new XElement(ns + "issn", value);
            el.SetAttributeValue("pub-type", pubType);
            root.Add(el);
            Console.WriteLine($"[UpdateMeta] Created <issn pub-type='{pubType}'> = '{value}'");
            created++;
        }

        // Local-name → DTO getter. Mirror of the read map but only for fields we
        // can safely write back. Boolean fields are encoded as "yes"/"no" to match
        // the upstream pipeline's convention.
        private static readonly Dictionary<string, Func<IPubMetaDto, string?>> _icoreWriteMap =
            new(StringComparer.OrdinalIgnoreCase)
            {
                ["icore-publisher"]           = d => d.Publisher,
                ["icore-xmltype"]             = d => d.XmlType,
                ["icore-dtd"]                 = d => d.Dtd,
                ["icore-isCopyEditing"]       = d => d.Copyedit.HasValue   ? (d.Copyedit.Value   ? "yes" : "no") : null,
                ["icore-isPreEditing"]        = d => d.PreEditing.HasValue ? (d.PreEditing.Value ? "yes" : "no") : null,
                ["icore-TestingEXE"]          = d => d.Testing.HasValue    ? (d.Testing.Value    ? "yes" : "no") : null,
                ["icore-CopyEditingTemplate"] = d => d.CeTemplate,
                ["icore-pagination-platform"] = d => d.PaginationPlatform,
                ["icore-doi"]                 = d => d.Doi,
                ["icore-copyrightstatement"]  = d => d.Copyrights,
                ["itracks-jobcardid"]         = d => d.JobCard,
                ["language"]                  = d => d.Language,
                ["country"]                   = d => d.Country == "None" ? "" : d.Country, // map UI "None" back to empty
                ["book-id"]                   = d => d.BookOrJournalId,
                ["book-title"]                = d => d.BookTitle,
                ["isbn"]                      = d => d.Isbn,
                ["year"]                      = d => d.Year,
                ["month"]                     = d => d.Month,
                ["journalcode"]               = d => d.JournalName,
                ["journaltitle"]              = d => d.JournalTitle,
                ["ChapterType"]               = d => d.DocumentType,
                ["ProcessType"]               = d => d.ProcessType,
                ["ArticleType"]               = d => d.ArticleType,
                ["article-id"]                = d => d.ArticleId,
                ["icore-publisher-imprint"]   = d => d.PublisherImprint,
            };

        // ══════════════════════════════════════════════════════════════════════
        // DOCX Pre-Processor: Convert TIFF/BMP/WMF/EMF -> PNG inside the ZIP
        // ══════════════════════════════════════════════════════════════════════
        private MemoryStream? ConvertUnsupportedImages(MemoryStream docxStream)
        {
            docxStream.Position = 0;

            // First pass: check if any unsupported images exist
            bool needsConversion = false;
            try
            {
                using var checkZip = new ZipArchive(docxStream, ZipArchiveMode.Read, leaveOpen: true);
                foreach (var entry in checkZip.Entries)
                {
                    if (!entry.FullName.StartsWith("word/media/")) continue;
                    var imgExt = Path.GetExtension(entry.Name).ToLowerInvariant();
                    if (imgExt is ".tiff" or ".tif" or ".bmp" or ".wmf" or ".emf")
                    {
                        needsConversion = true;
                        break;
                    }
                }
            }
            catch { return null; }

            if (!needsConversion) return null;

            // Second pass: rebuild the DOCX with converted images
            docxStream.Position = 0;
            var outputStream = new MemoryStream();

            var entries = new List<(string fullName, byte[] data)>();
            var renames = new Dictionary<string, string>();

            using (var readZip = new ZipArchive(docxStream, ZipArchiveMode.Read, leaveOpen: true))
            {
                foreach (var entry in readZip.Entries)
                {
                    using var entryStream = entry.Open();
                    using var entryMs = new MemoryStream();
                    entryStream.CopyTo(entryMs);
                    var data = entryMs.ToArray();

                    if (entry.FullName.StartsWith("word/media/"))
                    {
                        var imgExt = Path.GetExtension(entry.Name).ToLowerInvariant();
                        if (imgExt is ".tiff" or ".tif" or ".bmp" or ".wmf" or ".emf")
                        {
                            Console.WriteLine($"[ImageConvert] Converting {entry.FullName} ({data.Length} bytes, {imgExt})...");

                            byte[]? pngData = ConvertImageToPng(data, imgExt);

                            if (pngData != null)
                            {
                                var newName = Path.ChangeExtension(entry.FullName, ".jpg");
                                renames[entry.FullName] = newName;
                                Console.WriteLine($"[ImageConvert]   OK -> {newName} ({pngData.Length} bytes)");
                                entries.Add((newName, pngData));
                                continue;
                            }
                            else
                            {
                                Console.WriteLine($"[ImageConvert]   FAILED - keeping original");
                            }
                        }
                    }

                    entries.Add((entry.FullName, data));
                }
            }

            if (renames.Count == 0)
            {
                outputStream.Dispose();
                return null;
            }

            // Fix XML references in [Content_Types].xml and .rels files
            for (int i = 0; i < entries.Count; i++)
            {
                var (name, data) = entries[i];

                if (name == "[Content_Types].xml" || name.EndsWith(".rels"))
                {
                    var xml = Encoding.UTF8.GetString(data);
                    bool changed = false;

                    foreach (var (oldPath, newPath) in renames)
                    {
                        var oldTarget = oldPath.Replace("word/", "");
                        var newTarget = newPath.Replace("word/", "");
                        if (xml.Contains(oldTarget))
                        {
                            xml = xml.Replace(oldTarget, newTarget);
                            changed = true;
                        }
                        if (xml.Contains(oldPath))
                        {
                            xml = xml.Replace(oldPath, newPath);
                            changed = true;
                        }
                        if (xml.Contains("/" + oldPath))
                        {
                            xml = xml.Replace("/" + oldPath, "/" + newPath);
                            changed = true;
                        }
                    }

                    if (name == "[Content_Types].xml")
                    {
                        xml = Regex.Replace(xml,
                            @"<Default\s+Extension=""tiff?""\s+ContentType=""image/tiff""\s*/>",
                            "", RegexOptions.IgnoreCase);
                        xml = Regex.Replace(xml,
                            @"<Default\s+Extension=""bmp""\s+ContentType=""image/(bmp|x-bmp)""\s*/>",
                            "", RegexOptions.IgnoreCase);
                        xml = Regex.Replace(xml,
                            @"<Default\s+Extension=""[we]mf""\s+ContentType=""image/x-[we]mf""\s*/>",
                            "", RegexOptions.IgnoreCase);

                        xml = xml.Replace("image/tiff", "image/jpeg");
                        xml = xml.Replace("image/bmp", "image/jpeg");
                        xml = xml.Replace("image/x-wmf", "image/jpeg");
                        xml = xml.Replace("image/x-emf", "image/jpeg");

                        if (!xml.Contains(@"Extension=""jpg""") && !xml.Contains(@"Extension=""jpeg"""))
                        {
                            xml = xml.Replace("</Types>",
                                @"<Default Extension=""jpg"" ContentType=""image/jpeg""/></Types>");
                        }

                        changed = true;
                    }

                    if (changed)
                    {
                        entries[i] = (name, Encoding.UTF8.GetBytes(xml));
                    }
                }
            }

            // Write the new ZIP
            using (var writeZip = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var (name, data) in entries)
                {
                    var newEntry = writeZip.CreateEntry(name, CompressionLevel.Fastest);
                    using var newStream = newEntry.Open();
                    newStream.Write(data, 0, data.Length);
                }
            }

            Console.WriteLine($"[ImageConvert] Converted {renames.Count} images. New DOCX: {outputStream.Length} bytes");

            outputStream.Position = 0;
            return outputStream;
        }

        private byte[]? ConvertImageToPng(byte[] imageData, string sourceExt)
        {
            // Try System.Drawing first (Windows — handles TIFF/WMF/EMF reliably)
            try
            {
                #pragma warning disable CA1416
                using var imgStream = new MemoryStream(imageData);
                using var original = System.Drawing.Image.FromStream(imgStream);

                int w = original.Width, h = original.Height;
                if (w > MaxImageDimension || h > MaxImageDimension)
                {
                    double scale = Math.Min((double)MaxImageDimension / w, (double)MaxImageDimension / h);
                    w = (int)(w * scale);
                    h = (int)(h * scale);
                }
                using var resized = new System.Drawing.Bitmap(w, h);
                using (var g = System.Drawing.Graphics.FromImage(resized))
                {
                    g.Clear(System.Drawing.Color.White);
                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                    g.DrawImage(original, 0, 0, w, h);
                }

                using var outStream = new MemoryStream();
                var jpegCodec = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
                    .First(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);
                var encoderParams = new System.Drawing.Imaging.EncoderParameters(1);
                encoderParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
                    System.Drawing.Imaging.Encoder.Quality, JpegQuality);
                resized.Save(outStream, jpegCodec, encoderParams);

                Console.WriteLine($"[ImageConvert]   System.Drawing: {original.Width}x{original.Height} -> {w}x{h} JPEG ({outStream.Length / 1024}KB)");
                return outStream.ToArray();
                #pragma warning restore CA1416
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ImageConvert]   System.Drawing failed: {ex.Message}");
            }

            // SkiaSharp fallback (Linux/macOS/mobile)
            try
            {
                using var skBitmap = SkiaSharp.SKBitmap.Decode(imageData);
                if (skBitmap != null && skBitmap.Width > 0)
                {
                    int w = skBitmap.Width, h = skBitmap.Height;
                    SkiaSharp.SKBitmap target = skBitmap;

                    if (w > MaxImageDimension || h > MaxImageDimension)
                    {
                        double scale = Math.Min((double)MaxImageDimension / w, (double)MaxImageDimension / h);
                        w = (int)(w * scale); h = (int)(h * scale);
                        target = skBitmap.Resize(new SkiaSharp.SKImageInfo(w, h), SkiaSharp.SKFilterQuality.High);
                    }

                    using var skImage = SkiaSharp.SKImage.FromBitmap(target);
                    using var jpegData = skImage.Encode(SkiaSharp.SKEncodedImageFormat.Jpeg, (int)JpegQuality);
                    if (target != skBitmap) target.Dispose();

                    if (jpegData != null && jpegData.Size > 0)
                    {
                        Console.WriteLine($"[ImageConvert]   SkiaSharp: {skBitmap.Width}x{skBitmap.Height} -> {w}x{h} JPEG ({jpegData.Size / 1024}KB)");
                        return jpegData.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ImageConvert]   SkiaSharp failed: {ex.Message}");
            }

            return null;
        }

        // ── Helpers ──────────────────────────────────────────────────────────
        public static FormatType GetFormatType(string? fileName) =>
            Path.GetExtension(fileName ?? "").ToLowerInvariant() switch
            {
                ".docx" => FormatType.Docx,
                ".doc" => FormatType.Doc,
                ".rtf" => FormatType.Rtf,
                ".txt" => FormatType.Txt,
                _ => FormatType.Docx
            };

        public static string GetMimeType(string fileName) =>
            Path.GetExtension(fileName).ToLowerInvariant() switch
            {
                ".docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                ".doc" => "application/msword",
                ".rtf" => "application/rtf",
                ".txt" => "text/plain",
                ".xml" => "application/xml",
                _ => "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            };

        private static FormatType GetClipboardFormat(string? type) =>
            type?.ToLowerInvariant() switch
            {
                "html" => FormatType.Html,
                "rtf" => FormatType.Rtf,
                _ => FormatType.Html
            };
    }
}
