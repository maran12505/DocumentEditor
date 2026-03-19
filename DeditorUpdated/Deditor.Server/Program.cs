using Microsoft.AspNetCore.ResponseCompression;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllersWithViews();
builder.Services.AddHttpClient();   // for IHttpClientFactory in AiController
builder.Services.AddResponseCompression(opts =>
{
    opts.MimeTypes = ResponseCompressionDefaults.MimeTypes.Concat(
        new[] { "application/octet-stream" });
});

// Allow large uploads (manuscripts, books)
builder.Services.Configure<Microsoft.AspNetCore.Http.Features.FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 104857600; // 100 MB
});

var app = builder.Build();

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
app.MapControllers();
app.MapFallbackToFile("index.html");

app.Run();