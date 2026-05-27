// export-images.cs - PowerPoint COM を使って PPTX を JPG にエクスポートする（Visual QA 用）
// 使用例: dotnet run .github/skills/pptx-slide/assets/scripts/export-images.cs <input.pptx> [output-dir] [width=1280] [height=720]
//   - input.pptx       : エクスポート対象の PPTX
//   - output-dir       : JPG 出力先（省略時は input と同じディレクトリ）
//   - width/height     : 画像サイズ（省略時は 1280x720）
//
// 出力ファイル名: slide-01.jpg, slide-02.jpg, ...
//
// 前提: Windows + PowerPoint（COM）

#:property TargetFramework=net10.0-windows
#:property BuiltInComInteropSupport=true

using System.Reflection;
using System.Runtime.InteropServices;

if (args.Length < 1)
{
    Console.Error.WriteLine("Usage: dotnet run export-images.cs <input.pptx> [output-dir] [width] [height]");
    return 1;
}

Console.OutputEncoding = System.Text.Encoding.UTF8;

var inputPath = Path.GetFullPath(args[0]);
if (!File.Exists(inputPath))
{
    Console.Error.WriteLine($"Error: File not found: {inputPath}");
    return 1;
}

var outputDir = args.Length >= 2 ? Path.GetFullPath(args[1]) : Path.GetDirectoryName(inputPath)!;
int width  = args.Length >= 3 ? int.Parse(args[2]) : 1280;
int height = args.Length >= 4 ? int.Parse(args[3]) : 720;

if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine("Error: This script requires Windows + PowerPoint.");
    return 1;
}

Type? pptType = Type.GetTypeFromProgID("PowerPoint.Application");
if (pptType == null)
{
    Console.Error.WriteLine("Error: PowerPoint is not installed.");
    return 1;
}

object? app = null;
object? presObj = null;
try
{
    app = Activator.CreateInstance(pptType)!;

    // PowerPoint は Visible を msoTrue (-1) にしないと開けない。
    // dynamic 経由だと VARIANT 変換が安定しないため、すべて InvokeMember を使う。
    Invoke(app, "Visible", BindingFlags.SetProperty, -1);

    var presentations = Invoke(app, "Presentations", BindingFlags.GetProperty)!;
    // Open(FileName, ReadOnly, Untitled, WithWindow)
    presObj = Invoke(presentations, "Open", BindingFlags.InvokeMethod,
        inputPath, -1 /*ReadOnly=msoTrue*/, 0 /*Untitled=msoFalse*/, -1 /*WithWindow=msoTrue*/)!;

    var slides = Invoke(presObj, "Slides", BindingFlags.GetProperty)!;
    int slideCount = (int)Invoke(slides, "Count", BindingFlags.GetProperty)!;
    Console.WriteLine($"Slides: {slideCount}");

    for (int i = 1; i <= slideCount; i++)
    {
        var slide = Invoke(slides, "Item", BindingFlags.InvokeMethod, i)!;
        var outPath = Path.Combine(outputDir, $"slide-{i:D2}.jpg");
        // Slide.Export(FileName, FilterName, ScaleWidth, ScaleHeight)
        Invoke(slide, "Export", BindingFlags.InvokeMethod, outPath, "JPG", width, height);
        Marshal.FinalReleaseComObject(slide);
        Console.WriteLine($"  Exported: {outPath}");
    }

    Marshal.FinalReleaseComObject(slides);
    Console.WriteLine("✅ Export complete");
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    return 1;
}
finally
{
    try { if (presObj != null) Invoke(presObj, "Close", BindingFlags.InvokeMethod); } catch { }
    try { if (app != null) Invoke(app, "Quit", BindingFlags.InvokeMethod); } catch { }
    if (presObj != null) Marshal.FinalReleaseComObject(presObj);
    if (app != null) Marshal.FinalReleaseComObject(app);
}

static object? Invoke(object target, string name, BindingFlags flags, params object?[] args)
{
    return target.GetType().InvokeMember(name, flags, null, target, args);
}
