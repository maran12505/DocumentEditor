using Deditor.Core.Services.Interfaces;

namespace Deditor.Web.Services
{
    /// <summary>Web: same origin, no base URL needed.</summary>
    public class WebAppConfig : IAppConfig
    {
        public string ServerBaseUrl => "";
    }
}
