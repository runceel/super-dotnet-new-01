// list-layouts.cs - PPTX 内のスライドレイアウト一覧を取得する
// 使用例: dotnet run .github/skills/pptx-slide/assets/scripts/list-layouts.cs <pptx-path>
//         （引数省略時はスキル同梱の seed.pptx を読む）

using System.IO.Compression;
using System.Text.RegularExpressions;

Console.OutputEncoding = System.Text.Encoding.UTF8;

string seedPath;
if (args.Length >= 1)
{
    seedPath = Path.GetFullPath(args[0]);
}
else
{
    // スクリプトファイル自身の場所から seed.pptx を解決
    var scriptDir = Path.GetDirectoryName(typeof(Program).Assembly.Location)!;
    // file-based app では assembly location が一時パスなので、引数を必須にする
    Console.Error.WriteLine("Usage: dotnet run list-layouts.cs <pptx-path>");
    return 1;
}

if (!File.Exists(seedPath))
{
    Console.Error.WriteLine($"Error: File not found: {seedPath}");
    return 1;
}

using var zip = ZipFile.OpenRead(seedPath);

var layouts = zip.Entries
    .Where(e => e.FullName.StartsWith("ppt/slideLayouts/slideLayout") && e.FullName.EndsWith(".xml"))
    .OrderBy(e => {
        var m = Regex.Match(e.FullName, @"slideLayout(\d+)\.xml");
        return m.Success ? int.Parse(m.Groups[1].Value) : 0;
    })
    .ToList();

Console.WriteLine($"Total layouts: {layouts.Count}");
foreach (var entry in layouts)
{
    using var stream = entry.Open();
    using var reader = new StreamReader(stream);
    var content = reader.ReadToEnd();
    var nameMatch = Regex.Match(content, @"<p:cSld[^>]*name=""([^""]+)""");
    var name = nameMatch.Success ? nameMatch.Groups[1].Value : "(no name)";
    var num = Regex.Match(entry.FullName, @"slideLayout(\d+)").Groups[1].Value;
    Console.WriteLine($"  {num,3}: {name}");
}
return 0;
