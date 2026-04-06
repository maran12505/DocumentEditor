using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Deditor.Core.Services;
using Deditor.Core.Services.Interfaces;
using Deditor.Maui.Services;

namespace Deditor.Maui
{
    public static class MauiProgram
    {
        public static MauiApp CreateMauiApp()
        {
            // Register Syncfusion license
            Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense(
                "ORg4AjUWIQA/Gnt3VVhhQlJDfV5AQmBIYVp/TGpJfl96cVxMZVVBJAtUQF1hTH5bdE1jUH9ddX1VQWhbWkdy");

            var builder = MauiApp.CreateBuilder();
            builder
                .UseMauiApp<App>()
                .ConfigureFonts(fonts =>
                {
                    fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular");
                });

            builder.Services.AddMauiBlazorWebView();

#if DEBUG
            builder.Services.AddBlazorWebViewDeveloperTools();
            builder.Logging.AddDebug();
#endif

            // Load configuration from bundled appsettings.json
            var config = BuildConfiguration();
            builder.Services.AddSingleton<IConfiguration>(config);

            // HttpClient pointing to the DEditor server for API calls
            builder.Services.AddScoped(sp => new HttpClient
            {
                BaseAddress = new Uri("http://localhost:5278")
            });

            // HttpClient factory for external API calls (OpenRouter, Claude)
            builder.Services.AddHttpClient();

            // Platform config — server base URL for JS fetch calls
            builder.Services.AddSingleton<IAppConfig, MauiAppConfig>();

            // Register Core services DIRECTLY — runs locally, no HTTP server needed
            builder.Services.AddSingleton<IDocumentService, DocumentService>();
            builder.Services.AddSingleton<IAiService, AiService>();
            builder.Services.AddSingleton<IManuscriptService, ManuscriptService>();

            return builder.Build();
        }

        private static IConfiguration BuildConfiguration()
        {
            // Load appsettings.json from MAUI raw resources
            var configBuilder = new ConfigurationBuilder();

            try
            {
                using var stream = FileSystem.OpenAppPackageFileAsync("appsettings.json").Result;
                if (stream != null)
                {
                    // Copy to a MemoryStream since ConfigurationBuilder needs a seekable stream
                    var ms = new MemoryStream();
                    stream.CopyTo(ms);
                    ms.Position = 0;
                    configBuilder.AddJsonStream(ms);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[MAUI] Could not load appsettings.json: {ex.Message}");
            }

            return configBuilder.Build();
        }
    }
}
