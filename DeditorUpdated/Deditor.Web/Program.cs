using Microsoft.AspNetCore.Components.Web;
using Microsoft.AspNetCore.Components.WebAssembly.Hosting;
using Microsoft.Extensions.DependencyInjection;
using Deditor.Shared.UI;
using Deditor.Core.Services.Interfaces;
using Deditor.Web.Services;
using Syncfusion.Blazor;

// Register Syncfusion license — must be before builder.Build()
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("ORg4AjUWIQA/Gnt3VVhhQlJDfV5AQmBIYVp/TGpJfl96cVxMZVVBJAtUQF1hTH5bdE1jUH9ddX1VQWhbWkdy");

var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");

// HttpClient points to the server (same origin when hosted)
builder.Services.AddScoped(sp => new HttpClient
{
    BaseAddress = new Uri(builder.HostEnvironment.BaseAddress)
});

// Platform config — empty base URL for web (same origin)
builder.Services.AddScoped<IAppConfig, WebAppConfig>();

// Syncfusion Blazor components (DropDowns / Inputs / Buttons used by the iPubEdit popup).
// The sfBlazor JS bundle is loaded explicitly in index.html so it's ready before first render
// (avoids a WASM race where components render before the dynamically-injected script is parsed).
builder.Services.AddSyncfusionBlazor();

// Register HTTP-based service implementations (call server API over HTTP)
builder.Services.AddScoped<IDocumentService, HttpDocumentService>();
builder.Services.AddScoped<IAiService, HttpAiService>();
builder.Services.AddScoped<IManuscriptService, HttpManuscriptService>();

await builder.Build().RunAsync();
