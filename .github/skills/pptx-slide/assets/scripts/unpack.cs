// unpack.cs - PPTX を展開して XML を pretty-print する
// 使用例: dotnet run .github/skills/pptx-slide/assets/scripts/unpack.cs <input.pptx> <output-dir>

using System.IO.Compression;
using System.Text;
using System.Xml;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: dotnet run temp/scripts/unpack.cs <input.pptx> <output-dir>");
    return 1;
}

var inputPath = Path.GetFullPath(args[0]);
var outputDir = Path.GetFullPath(args[1]);

if (!File.Exists(inputPath))
{
    Console.Error.WriteLine($"Error: File not found: {inputPath}");
    return 1;
}

if (Directory.Exists(outputDir))
{
    Console.WriteLine($"Removing existing directory: {outputDir}");
    Directory.Delete(outputDir, recursive: true);
}

Directory.CreateDirectory(outputDir);

using var zip = ZipFile.OpenRead(inputPath);

foreach (var entry in zip.Entries)
{
    var destPath = Path.Combine(outputDir, entry.FullName.Replace('/', Path.DirectorySeparatorChar));

    if (entry.FullName.EndsWith('/'))
    {
        Directory.CreateDirectory(destPath);
        continue;
    }

    var dir = Path.GetDirectoryName(destPath)!;
    if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

    if (entry.FullName.EndsWith(".xml") || entry.FullName.EndsWith(".rels"))
    {
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8);
        var content = reader.ReadToEnd();
        var pretty = PrettyPrintXml(content);
        File.WriteAllText(destPath, pretty, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }
    else
    {
        entry.ExtractToFile(destPath, overwrite: true);
    }
}

Console.WriteLine($"✅ Unpacked {zip.Entries.Count} entries to: {outputDir}");
return 0;

static string PrettyPrintXml(string xml)
{
    try
    {
        var doc = new XmlDocument();
        doc.LoadXml(xml);

        // StringBuilderだとUTF-16になるのでMemoryStreamを使ってUTF-8を保持する
        using var ms = new MemoryStream();
        var settings = new XmlWriterSettings
        {
            Indent = true,
            IndentChars = "  ",
            NewLineChars = "\n",
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            OmitXmlDeclaration = false
        };

        using (var writer = XmlWriter.Create(ms, settings))
        {
            doc.Save(writer);
        }
        return new UTF8Encoding(false).GetString(ms.ToArray());
    }
    catch
    {
        return xml;
    }
}
