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
        Stream Save(string sfdtContent, string fileName);

        /// <summary>
        /// Convert system clipboard content (HTML/RTF) to SFDT for pasting into the editor.
        /// </summary>
        string SystemClipboard(string content, string type);

        /// <summary>
        /// Process a metafile/image service request from the Syncfusion editor.
        /// </summary>
        string ServiceBase(string? imageData, string? action);
    }
}
