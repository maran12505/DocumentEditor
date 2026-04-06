# DEditor Restructure Plan: Web + Native from a Shared Codebase

## Goal
Restructure from the current 2-project layout (Client + Server) into a **5-project architecture** that shares all UI and business logic between Blazor WASM (web) and MAUI Blazor Hybrid (desktop/mobile), with offline-capable local processing in MAUI.

---

## Current Structure
```
Deditor.sln
├── Deditor.Client   (Blazor WASM)  - All UI, pages, JS, CSS
└── Deditor.Server   (ASP.NET Core) - API controllers, doc processing, AI calls
```

## Target Structure
```
Deditor.sln
│
├── Deditor.Core/                        ← NEW: Shared business logic library
│   ├── Deditor.Core.csproj              (net10.0, no UI dependencies)
│   ├── Models/
│   │   ├── SaveParameter.cs
│   │   ├── ChatRequest.cs
│   │   ├── EssentialCapsRequest.cs
│   │   ├── ManuscriptRequest.cs
│   │   ├── ManuscriptResponse.cs
│   │   ├── ChangeItem.cs
│   │   ├── ChangeResponse.cs
│   │   ├── CustomParameter.cs
│   │   └── MetafileParam.cs
│   └── Services/
│       ├── Interfaces/
│       │   ├── IDocumentService.cs      (Import, Save, SystemClipboard, SpellCheck)
│       │   ├── IAiService.cs            (Chat, EssentialCaps)
│       │   └── IManuscriptService.cs    (Process)
│       ├── DocumentService.cs           (extracted from DocumentEditorController)
│       ├── AiService.cs                 (extracted from AiController)
│       └── ManuscriptService.cs         (extracted from ManuscriptProcessorController)
│
├── Deditor.Shared.UI/                   ← NEW: Razor Class Library (all shared UI)
│   ├── Deditor.Shared.UI.csproj         (net10.0, Microsoft.NET.Sdk.Razor)
│   ├── App.razor                        (moved from Client)
│   ├── _Imports.razor                   (updated namespaces)
│   ├── Layout/
│   │   └── MainLayout.razor             (moved from Client)
│   ├── Pages/
│   │   ├── DocumentEditor.razor         (moved from Client, @inject IDocumentService)
│   │   ├── DocumentEditor.razor.css
│   │   └── Welcome.razor                (moved from Client)
│   └── wwwroot/                         (moved from Client)
│       ├── css/ (app.css, material.css, fontdialog.css, etc.)
│       ├── js/  (ej2.min.js, documenteditor.js, compromise.js, etc.)
│       └── images/ (Logo.png, Logo1.png, Logo.jpg)
│
├── Deditor.Web/                         ← RENAMED from Deditor.Client (thin WASM host)
│   ├── Deditor.Web.csproj               (BlazorWebAssembly SDK)
│   ├── Program.cs                       (registers HttpClient-based service stubs)
│   ├── Services/
│   │   ├── HttpDocumentService.cs       (IDocumentService → calls /api/documenteditor/*)
│   │   ├── HttpAiService.cs             (IAiService → calls /api/ai/*)
│   │   └── HttpManuscriptService.cs     (IManuscriptService → calls /api/manuscript/*)
│   └── wwwroot/
│       └── index.html                   (WASM-specific: loads blazor.webassembly.js)
│
├── Deditor.Server/                      ← SIMPLIFIED (thin API shell)
│   ├── Deditor.Server.csproj            (references Core, not Client directly)
│   ├── Program.cs                       (hosts WASM + registers Core services)
│   ├── Controllers/
│   │   ├── DocumentEditorController.cs  (thin: delegates to IDocumentService)
│   │   ├── AiController.cs              (thin: delegates to IAiService)
│   │   └── ManuscriptProcessorController.cs (thin: delegates to IManuscriptService)
│   └── appsettings.json
│
└── Deditor.Maui/                        ← NEW: MAUI Blazor Hybrid app
    ├── Deditor.Maui.csproj              (MAUI multi-target: Win/Mac/Android/iOS)
    ├── MauiProgram.cs                   (registers LOCAL Core services directly)
    ├── MainPage.xaml                    (BlazorWebView host)
    ├── MainPage.xaml.cs
    ├── appsettings.json                 (bundled as raw asset for API keys)
    ├── wwwroot/
    │   └── index.html                   (MAUI-specific: loads blazor.webview.js)
    ├── Platforms/
    │   ├── Android/
    │   ├── iOS/
    │   ├── MacCatalyst/
    │   └── Windows/
    └── Resources/
        └── AppIcon, Splash, etc.
```

---

## Implementation Steps

### Step 1: Create `Deditor.Core` (Shared Business Logic)

**1a. Create project & models**
- Create `Deditor.Core.csproj` targeting `net10.0`
- Add NuGet packages: `DocumentFormat.OpenXml`, `Syncfusion.DocIO.Net.Core`, `Syncfusion.EJ2.WordEditor.AspNet.Core`, `SkiaSharp`, `System.Drawing.Common`, `Newtonsoft.Json`
- Extract all DTO/model classes from the 3 controllers into `Models/`:
  - `SaveParameter`, `CustomParameter`, `RestrictParameter`, `SpellCheckParameter`, `MetafileParam`
  - `ChatRequest`, `EssentialCapsRequest`
  - `ManuscriptRequest`, `ManuscriptResponse`, `ChangeItem`, `ChangeResponse`

**1b. Define service interfaces**
- `IDocumentService`:
  ```csharp
  Task<string> ImportAsync(Stream fileStream, string fileName, long fileLength);
  Stream Save(string sfdtContent, string fileName);
  string SystemClipboard(string content, string type);
  ```
- `IAiService`:
  ```csharp
  Task<string> ChatAsync(string prompt);
  Task<string> EssentialCapsAsync(string text);
  ```
- `IManuscriptService`:
  ```csharp
  Task<ManuscriptResponse> ProcessAsync(string sfdt);
  ```

**1c. Extract business logic into service implementations**
- `DocumentService.cs`: Move ALL logic from `DocumentEditorController` — import pipeline, `ConvertUnsupportedImages()`, `ConvertImageToPng()`, format helpers, caching
- `AiService.cs`: Move chat/proofreading logic from `AiController` — takes `IHttpClientFactory` + `IConfiguration` via constructor DI
- `ManuscriptService.cs`: Move manuscript processing from `ManuscriptProcessorController` — `ExtractTextWithIndex()`, `ApplyTrackedChanges()`, `CallClaudeApi()`, `GenerateMockChanges()`, all helpers

### Step 2: Create `Deditor.Shared.UI` (Razor Class Library)

**2a. Create RCL project**
- Create `Deditor.Shared.UI.csproj` with `Microsoft.NET.Sdk.Razor`, targeting `net10.0`
- Add package refs: `Microsoft.AspNetCore.Components.Web`, `Syncfusion.Licensing`
- Add project reference to `Deditor.Core`

**2b. Move all UI from Client**
- Move `App.razor`, `_Imports.razor`, `Layout/MainLayout.razor`
- Move `Pages/DocumentEditor.razor` + `.razor.css`, `Pages/Welcome.razor`
- Delete `Pages/DE.razor` (dropping per user request)
- Move entire `wwwroot/` (css/, js/, images/) — BUT NOT `index.html` (each host gets its own)

**2c. Update Razor component code**
- Replace `@inject HttpClient Http` with `@inject IDocumentService DocService` (and IAiService, IManuscriptService)
- Replace all `Http.PostAsync("api/documenteditor/import", ...)` → `DocService.ImportAsync(...)`
- Replace `Http.PostAsJsonAsync("api/documenteditor/save", ...)` → `DocService.Save(...)`
- Replace `Http.PostAsJsonAsync("api/ai/chat", ...)` → `AiService.ChatAsync(...)`
- Replace `Http.PostAsJsonAsync("api/manuscript/process", ...)` → `ManuscriptService.ProcessAsync(...)`
- Update `_Imports.razor` namespaces: `Deditor.Client` → `Deditor.Shared.UI`, add `@using Deditor.Core.Services.Interfaces`, `@using Deditor.Core.Models`

### Step 3: Create `Deditor.Web` (Thin WASM Host)

**3a. Create WASM project**
- Create `Deditor.Web.csproj` using `Microsoft.NET.Sdk.BlazorWebAssembly`
- Add package refs: WASM packages, Syncfusion.Licensing
- Add project references to `Deditor.Shared.UI` and `Deditor.Core`

**3b. HTTP-based service implementations**
- `HttpDocumentService.cs` — implements `IDocumentService` by calling `/api/documenteditor/*` via `HttpClient` (same HTTP calls as current DocumentEditor.razor)
- `HttpAiService.cs` — implements `IAiService` by calling `/api/ai/*`
- `HttpManuscriptService.cs` — implements `IManuscriptService` by calling `/api/manuscript/*`
- These are simple pass-through classes that serialize/deserialize JSON over HTTP

**3c. Program.cs**
```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("...");
var builder = WebAssemblyHostBuilder.CreateDefault(args);
builder.RootComponents.Add<App>("#app");
builder.RootComponents.Add<HeadOutlet>("head::after");
builder.Services.AddScoped(sp => new HttpClient { BaseAddress = new Uri(builder.HostEnvironment.BaseAddress) });
builder.Services.AddScoped<IDocumentService, HttpDocumentService>();
builder.Services.AddScoped<IAiService, HttpAiService>();
builder.Services.AddScoped<IManuscriptService, HttpManuscriptService>();
await builder.Build().RunAsync();
```

**3d. wwwroot/index.html**
- Copy current `index.html` from Client
- Change static asset paths: `css/app.css` → `_content/Deditor.Shared.UI/css/app.css` (RCL static asset convention)
- Same for all JS/CSS refs: prefix with `_content/Deditor.Shared.UI/`
- Keep `_framework/blazor.webassembly.js` as the runtime script

### Step 4: Simplify `Deditor.Server`

**4a. Update project references**
- Remove direct reference to old `Deditor.Client`
- Add references to `Deditor.Core` and `Deditor.Web`
- Remove heavy NuGet packages that moved to Core (Syncfusion, OpenXML, SkiaSharp)

**4b. Thin out controllers**
- Each controller becomes a thin HTTP adapter that:
  - Accepts the HTTP request
  - Delegates to the corresponding `Deditor.Core` service
  - Returns the result
- Example `DocumentEditorController.Import()`:
  ```csharp
  [HttpPost("import")]
  public async Task<IActionResult> Import(IFormFile files) {
      using var ms = new MemoryStream();
      await files.CopyToAsync(ms);
      ms.Position = 0;
      var sfdt = await _documentService.ImportAsync(ms, files.FileName, files.Length);
      return Ok(sfdt);
  }
  ```

**4c. Update Program.cs**
- Register Core services via DI: `builder.Services.AddSingleton<IDocumentService, DocumentService>()` etc.
- Keep Kestrel limits, compression, static files middleware
- Change `UseBlazorFrameworkFiles` to reference `Deditor.Web`

### Step 5: Create `Deditor.Maui` (MAUI Blazor Hybrid)

**5a. Create MAUI project**
- Create `Deditor.Maui.csproj` using `Microsoft.NET.Sdk.Maui` with Blazor support
- Target frameworks: `net10.0-android`, `net10.0-ios`, `net10.0-maccatalyst`, `net10.0-windows10.0.19041.0`
- Add project references to `Deditor.Shared.UI` and `Deditor.Core`
- Add NuGet: `Microsoft.AspNetCore.Components.WebView.Maui`, SkiaSharp MAUI native assets

**5b. MauiProgram.cs**
```csharp
Syncfusion.Licensing.SyncfusionLicenseProvider.RegisterLicense("...");
var builder = MauiApp.CreateBuilder();
builder.UseMauiApp<App>()
       .ConfigureFonts(fonts => { fonts.AddFont("OpenSans-Regular.ttf", "OpenSansRegular"); });
builder.Services.AddMauiBlazorWebView();
#if DEBUG
builder.Services.AddBlazorWebViewDeveloperTools();
#endif
// Register Core services DIRECTLY — no HTTP, runs locally
builder.Services.AddSingleton<IDocumentService, DocumentService>();
builder.Services.AddSingleton<IAiService, AiService>();
builder.Services.AddSingleton<IManuscriptService, ManuscriptService>();
builder.Services.AddSingleton<IConfiguration>(BuildConfiguration());
builder.Services.AddHttpClient(); // needed for AI API calls
```

**5c. MainPage.xaml**
```xml
<ContentPage xmlns="http://schemas.microsoft.com/dotnet/2021/maui"
             xmlns:x="http://schemas.microsoft.com/winfx/2009/xaml"
             x:Class="Deditor.Maui.MainPage">
    <BlazorWebView HostPage="wwwroot/index.html">
        <BlazorWebView.RootComponents>
            <RootComponent Selector="#app" ComponentType="{x:Type shared:App}" />
        </BlazorWebView.RootComponents>
    </BlazorWebView>
</ContentPage>
```

**5d. wwwroot/index.html (MAUI-specific)**
- Similar to Web's index.html but:
  - Static assets from RCL: `_content/Deditor.Shared.UI/css/app.css` etc.
  - Runtime script: `_framework/blazor.webview.js` (NOT `blazor.webassembly.js`)
  - No SPA fallback needed (MAUI handles routing)

**5e. App configuration for MAUI**
- Bundle `appsettings.json` as a Raw asset in the MAUI project
- Load configuration at startup via `ConfigurationBuilder.AddJsonStream()`
- API keys (OpenRouter, Claude) stored in bundled config or MAUI SecureStorage

### Step 6: Update Solution File

Add all 5 projects to `Deditor.sln` with proper project GUIDs and solution folders:
```
Solution Items
├── Deditor.Core
├── Deditor.Shared.UI
├── Deditor.Web        (replaces old Deditor.Client)
├── Deditor.Server
└── Deditor.Maui
```

---

## Project Dependency Graph
```
Deditor.Core          ← No UI dependencies. Pure .NET 10 library.
    ↑
Deditor.Shared.UI     ← Razor Class Library. References Core for interfaces.
    ↑         ↑
Deditor.Web   Deditor.Maui    ← Host projects. Each provides DI registrations.
    ↑
Deditor.Server        ← References Core + Web. Thin API layer.
```

## Key Design Principles

1. **Interface-driven services**: Components inject `IDocumentService`, never `HttpClient` directly. The host project decides the implementation.
2. **RCL static assets**: All JS/CSS/images live in the RCL and are served via `_content/Deditor.Shared.UI/` convention automatically by both WASM and MAUI.
3. **Offline-capable**: MAUI's `DocumentService` runs Syncfusion/OpenXML/SkiaSharp locally. AI features (Chat, Manuscript) require network but degrade gracefully.
4. **No code duplication**: Business logic exists once in Core, UI exists once in Shared.UI. Host projects are <100 lines each.

## Files to Delete
- `Deditor.Client/Pages/DE.razor` (dropped per user preference)
- Old `Deditor.Client/` project (replaced by `Deditor.Web/` + `Deditor.Shared.UI/`)
