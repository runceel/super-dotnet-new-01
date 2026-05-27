// pack.cs - 展開済みディレクトリを PPTX に再パックする
// 使用例: dotnet run .github/skills/pptx-slide/assets/scripts/pack.cs <unpacked-dir> <output.pptx>

using System.IO.Compression;
using System.Text;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: dotnet run pack.cs <unpacked-dir> <output.pptx>");
    return 1;
}

var inputDir  = Path.GetFullPath(args[0]);
var outputPath = Path.GetFullPath(args[1]);

if (!Directory.Exists(inputDir))
{
    Console.Error.WriteLine($"Error: Directory not found: {inputDir}");
    return 1;
}

if (File.Exists(outputPath)) File.Delete(outputPath);

var files = Directory.GetFiles(inputDir, "*", SearchOption.AllDirectories);
int count = 0;

using (var zip = ZipFile.Open(outputPath, ZipArchiveMode.Create))
{
    // [Content_Types].xml を最初に追加（Office の要件）
    var contentTypesPath = Path.Combine(inputDir, "[Content_Types].xml");
    if (File.Exists(contentTypesPath))
    {
        AddEntry(zip, contentTypesPath, "[Content_Types].xml");
        count++;
    }

    foreach (var file in files.OrderBy(f => f))
    {
        var relativePath = Path.GetRelativePath(inputDir, file).Replace(Path.DirectorySeparatorChar, '/');
        if (relativePath == "[Content_Types].xml") continue; // 既に追加済み

        AddEntry(zip, file, relativePath);
        count++;
    }
}

Console.WriteLine($"✅ Packed {count} files to: {outputPath}");
return 0;

static void AddEntry(ZipArchive zip, string filePath, string entryName)
{
    var entry = zip.CreateEntry(entryName, CompressionLevel.Optimal);
    using var dest = entry.Open();
    using var src  = File.OpenRead(filePath);
    src.CopyTo(dest);
}
