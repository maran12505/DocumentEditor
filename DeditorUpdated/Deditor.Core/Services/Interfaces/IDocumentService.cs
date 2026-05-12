using Deditor.Core.Models;

namespace Deditor.Core.Services.Interfaces
{
    /// <summary>
    /// Document import, save, clipboard, and spellcheck operations.
    /// Web: implemented via HTTP calls to the server API.
    /// MAUI: implemented locally using Syncfusion/OpenXML/SkiaSharp.
    /// </summary>
    public interface IDocumentService
    {
        /// <summary>
        /// Import a document file (DOCX, DOC, RTF, TXT, XML) and return its SFDT representation.
        /// Handles image conversion (TIFF/BMP/WMF/EMF to PNG) automatically.
        /// </summary>
        Task<string> ImportAsync(Stream fileStream, string fileName, long fileLength);

        /// <summary>
        /// Save SFDT content to a document stream in the format determined by the file extension.
        /// </summary>
        Task<Stream> SaveAsync(string sfdtContent, string fileName);

        /// <summary>
        /// Convert system clipboard content (HTML/RTF) to SFDT for pasting into the editor.
        /// </summary>
        Task<string> SystemClipboardAsync(string content, string type);

        /// <summary>
        /// Process a metafile/image service request from the Syncfusion editor.
        /// </summary>
        string ServiceBase(string? imageData, string? action);

        /// <summary>
        /// Extract iPubEdit metadata from the customXml parts of a DOCX file.
        /// Returns null if the file has no customXml or no recognized fields.
        /// Best-effort: any extraction error is swallowed and returns null.
        /// </summary>
        Task<IPubMetaDto?> ExtractIPubMetaAsync(Stream docxStream);

        /// <summary>
        /// Write iPubEdit metadata back into the customXml parts of a DOCX file.
        /// Updates the iCore part (root element local-name = "icore"), preferring
        /// /customXml/item3.xml when multiple parts exist. Existing elements get
        /// new text; missing elements are appended to the root. Original DOCX is
        /// returned unchanged on failure or when no iCore part is found.
        /// </summary>
        Task<byte[]> UpdateIPubCustomXmlAsync(byte[] docxBytes, IPubMetaDto dto);
    }
}
