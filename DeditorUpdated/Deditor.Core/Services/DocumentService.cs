using Deditor.Core.Services.Interfaces;
using Syncfusion.EJ2.DocumentEditor;
using Newtonsoft.Json;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

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
