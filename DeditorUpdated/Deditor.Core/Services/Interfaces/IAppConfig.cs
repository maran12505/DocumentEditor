namespace Deditor.Core.Services.Interfaces
{
    /// <summary>
    /// Platform-specific app configuration.
    /// Web: server base URL is empty (same origin).
    /// MAUI: server base URL points to the running server.
    /// </summary>
    public interface IAppConfig
    {
        /// <summary>
        /// Base URL of the DEditor API server.
        /// Empty string for Web (same origin), "http://localhost:5278" for MAUI.
        /// </summary>
        string ServerBaseUrl { get; }
    }
}
