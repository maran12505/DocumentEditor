namespace Deditor.Core.Services.Interfaces
{
    /// <summary>
    /// AI-powered copyediting, proofreading, and text transformation.
    /// Web: implemented via HTTP calls to the server API.
    /// MAUI: implemented locally by calling OpenRouter API directly.
    /// </summary>
    public interface IAiService
    {
        /// <summary>
        /// Send a copyediting/proofreading prompt and return the AI response.
        /// </summary>
        Task<string> ChatAsync(string prompt);

        /// <summary>
        /// Apply Essential Caps formatting to the given text.
        /// </summary>
        Task<string> EssentialCapsAsync(string text);
    }
}
