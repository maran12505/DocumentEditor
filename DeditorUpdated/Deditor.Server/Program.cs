using Microsoft.AspNetCore.ResponseCompression;
using Deditor.Core.Services;
using Deditor.Core.Services.Interfaces;

Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("ORg4AjUWIQA/Gnt3VVhhQlJDfV5AQmBIYVp/TGpJfl96cVxMZVVBJAtUQF1hTH5bdE1jUH9ddX1VQWhbWkdy");

var builder = WebApplication.CreateBuilder(args);

builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = 209_715_200;
    options.Limits.KeepAliveTimeout = TimeSpan.FromMinutes(5);
    options.Limits.RequestHeadersTimeout = TimeSpan.FromMinutes(5);
});

builder.Services.AddControllersWithViews();
builder.Services.AddHttpClient();

// CORS — allow MAUI BlazorWebView (origin https://0.0.0.1) to call the API
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowMaui", policy =>
    {
        policy.WithOrigins("https://0.0.0.1", "http://0.0.0.1")
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

// Register Core services — shared between Server controllers and MAUI
builder.Services.AddSingleton<IDocumentService, DocumentService>();
builder.Services.AddSingleton<IAiService, AiService>();
builder.Services.AddSingleton<IManuscriptService, ManuscriptService>();

// Enable gzip + brotli compression for ALL text responses (critical for SFDT)
builder.Services.AddResponseCompression(opts =>
{
    opts.EnableForHttps = true;
    opts.Providers.Add<BrotliCompressionProvider>();
    opts.Providers.Add<GzipCompressionProvider>();
    opts.MimeTypes = ResponseCompressionDefaults.MimeTypes.Concat(new[]
    {
        "application/json",
        "text/plain",
        "application/octet-stream"
    });
});

builder.Services.Configure<BrotliCompressionProviderOptions>(opts =>
{
    opts.Level = System.IO.Compression.CompressionLevel.Fastest;
});

builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 209_715_200;
});

var app = builder.Build();

// Response compression MUST be before static files and routing
app.UseResponseCompression();

if (app.Environment.IsDevelopment())
    app.UseWebAssemblyDebugging();
else
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseBlazorFrameworkFiles();
app.UseStaticFiles();
app.UseRouting();

// CORS must be after routing, before endpoints
app.UseCors("AllowMaui");

app.MapControllers();
app.MapFallbackToFile("index.html");

app.Run();
