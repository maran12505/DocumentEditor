using Deditor.Core.Services.Interfaces;

namespace Deditor.Maui.Services
{
    /// <summary>MAUI: points JS fetch calls to the running server.</summary>
    public class MauiAppConfig : IAppConfig
    {
        public string ServerBaseUrl => "http://localhost:5278";
    }
}
