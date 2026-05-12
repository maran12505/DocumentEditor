using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using ClosedXML.Excel;

// Converts List_Elements_DB.xlsx into elements-book.json + elements-journal.json.
// Hierarchy comes from the tlParentID column. Tag string comes from the tag column.
// Run from repo root: dotnet run --project tools/excel-to-elements-json -- <xlsx> <out-dir>

var xlsxPath = args.Length > 0
    ? args[0]
    : @"D:\profile backup\Documents\List_Elements_DB.xlsx";
var outDir = args.Length > 1
    ? args[1]
    : @"DeditorUpdated\Deditor.Shared.UI\wwwroot\data";

if (!File.Exists(xlsxPath))
{
    Console.Error.WriteLine($"ERROR: xlsx not found: {xlsxPath}");
    return 1;
}
Directory.CreateDirectory(outDir);

using var wb = new XLWorkbook(xlsxPath);
foreach (var ws in wb.Worksheets)
    Console.WriteLine($"Sheet: {ws.Name}  rows={ws.LastRowUsed()?.RowNumber()}  cols={ws.LastColumnUsed()?.ColumnNumber()}");

WriteSheet(wb, "Book",    Path.Combine(outDir, "elements-book.json"));
WriteSheet(wb, "Journal", Path.Combine(outDir, "elements-journal.json"));
return 0;

static void WriteSheet(XLWorkbook wb, string sheetName, string outPath)
{
    var ws = wb.Worksheets.FirstOrDefault(s => string.Equals(s.Name, sheetName, StringComparison.OrdinalIgnoreCase));
    if (ws is null)
    {
        Console.Error.WriteLine($"  skipping {sheetName} — sheet missing");
        return;
    }

    var lastRow = ws.LastRowUsed()?.RowNumber() ?? 0;
    var lastCol = ws.LastColumnUsed()?.ColumnNumber() ?? 0;
    if (lastRow < 2 || lastCol < 1)
    {
        Console.Error.WriteLine($"  {sheetName}: empty");
        return;
    }

    // Map header → column index
    var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
    for (int c = 1; c <= lastCol; c++)
    {
        var h = ws.Cell(1, c).GetString().Trim();
        if (!string.IsNullOrEmpty(h) && !headers.ContainsKey(h))
            headers[h] = c;
    }
    int Col(params string[] names)
    {
        foreach (var n in names)
            if (headers.TryGetValue(n, out var idx)) return idx;
        return -1;
    }

    int idCol     = Col("tlUniqID", "UniqID", "id");
    int nameCol   = Col("Elements", "Element", "Name");
    int parentCol = Col("tlParentID", "ParentID");
    int tagCol    = Col("tag", "TagList", "Tag List", "Tag");
    int attrCol   = Col("tlAttributeID", "AttributeID", "Attributes");

    Console.WriteLine($"  {sheetName}: id={idCol} name={nameCol} parent={parentCol} tag={tagCol} attr={attrCol} (1-based; -1=missing)");

    if (idCol < 0 || nameCol < 0 || parentCol < 0)
    {
        Console.Error.WriteLine($"  {sheetName}: required columns missing — wrote header dump:");
        foreach (var kv in headers.OrderBy(k => k.Value))
            Console.Error.WriteLine($"    [{kv.Value}] {kv.Key}");
        return;
    }

    var nodes = new Dictionary<int, Node>();
    var rawAttrIds = new Dictionary<int, List<int>>();
    for (int r = 2; r <= lastRow; r++)
    {
        var idStr = ws.Cell(r, idCol).GetString().Trim();
        if (!int.TryParse(idStr, out var id) || id == 0) continue;

        var name   = ws.Cell(r, nameCol).GetString().Trim();
        var parentStr = ws.Cell(r, parentCol).GetString().Trim();
        int.TryParse(parentStr, out var parentId);
        var rawTag = tagCol > 0 ? ws.Cell(r, tagCol).GetString().Trim() : "";
        var attrStr = attrCol > 0 ? ws.Cell(r, attrCol).GetString().Trim() : "";

        nodes[id] = new Node
        {
            Id = id,
            Name = string.IsNullOrEmpty(name) ? StripTag(rawTag) : name,
            RawTag = rawTag,
            ParentId = parentId,
            Children = new List<Node>()
        };
        var attrIds = ParseAttrIds(attrStr);
        if (attrIds.Count > 0) rawAttrIds[id] = attrIds;
    }

    // Resolve attribute IDs → attribute names by looking up other nodes by id
    foreach (var (nodeId, attrIds) in rawAttrIds)
    {
        var resolved = new List<Attr>();
        foreach (var aid in attrIds)
        {
            if (nodes.TryGetValue(aid, out var refNode) && !string.IsNullOrEmpty(refNode.Name))
                resolved.Add(new Attr { Name = refNode.Name });
        }
        if (resolved.Count > 0)
            nodes[nodeId].Attributes = resolved;
    }

    // Build tree
    var roots = new List<Node>();
    foreach (var n in nodes.Values)
    {
        if (n.ParentId == 0 || !nodes.TryGetValue(n.ParentId, out var parent))
            roots.Add(n);
        else
            parent.Children.Add(n);
    }

    // Stable sort children by id (preserves Excel ordering)
    void SortRec(List<Node> list)
    {
        list.Sort((a, b) => a.Id.CompareTo(b.Id));
        foreach (var c in list) SortRec(c.Children);
    }
    roots.Sort((a, b) => a.Id.CompareTo(b.Id));
    foreach (var r in roots) SortRec(r.Children);

    var json = JsonSerializer.Serialize(roots, new JsonSerializerOptions
    {
        WriteIndented = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
    });

    File.WriteAllText(outPath, json);
    Console.WriteLine($"  → {outPath}  roots={roots.Count}  total={nodes.Count}  size={new FileInfo(outPath).Length / 1024}KB");
}

static string StripTag(string tag)
{
    if (string.IsNullOrEmpty(tag)) return "";
    var t = tag.Trim();
    if (t.StartsWith('#')) t = t[1..];
    var semi = t.IndexOf(';');
    if (semi >= 0) t = t[..semi];
    return t;
}

static List<int> ParseAttrIds(string s)
{
    var ids = new List<int>();
    if (string.IsNullOrWhiteSpace(s)) return ids;
    foreach (var raw in s.Split(new[] { ',', ';', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        if (int.TryParse(raw.Trim(), out var v) && v > 0) ids.Add(v);
    return ids;
}

class Node
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public string? RawTag { get; set; }
    public int ParentId { get; set; }
    public List<Attr>? Attributes { get; set; }
    public List<Node> Children { get; set; } = new();
}

class Attr
{
    public string Name { get; set; } = "";
}
