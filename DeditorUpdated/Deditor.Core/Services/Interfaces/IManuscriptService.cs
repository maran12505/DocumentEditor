using Deditor.Core.Models;

namespace Deditor.Core.Services.Interfaces
{
    /// <summary>
    /// Full manuscript analysis via Claude API with tracked changes output.
    /// Web: implemented via HTTP calls to the server API.
    /// MAUI: implemented locally by calling Claude API directly.
    /// </summary>
    public interface IManuscriptService
    {
        /// <summary>
        /// Analyze manuscript SFDT content, get AI suggestions, and return
        /// a new SFDT with tracked changes applied.
        /// </summary>
        Task<ManuscriptResponse> ProcessAsync(string sfdt);
    }
}
