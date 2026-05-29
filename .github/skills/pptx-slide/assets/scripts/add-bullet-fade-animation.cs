// add-bullet-fade-animation.cs — 指定スライドの本文に箇条書きフェードアニメーションを追加する
//
// 使用例: dotnet run .github/skills/pptx-slide/assets/scripts/add-bullet-fade-animation.cs <input.pptx> <output.pptx> <slide-index> ...
//   - input.pptx   : 入力 PPTX
//   - output.pptx  : 出力 PPTX（input と同じパスを指定して上書きしても OK）
//   - slide-index  : アニメーションを追加するスライド番号（1-based、複数指定可）
//
// 動作:
//   - 各スライドの "Content Placeholder*" という名前の Shape を対象
//   - 既存アニメーションがあればクリアしてから Fade(presetID=10) を msoAnimateTextByFirstLevel で追加
//   - トップレベル箇条書きごとに 1 クリック、ぶら下がり子要素は親と同時に表示される
//
// 前提:
//   - Windows + PowerPoint（COM 連携で使用）
//   - 対象 PPTX ファイルは PowerPoint で開かれていないこと（ファイルロック回避のため）
//
// 詳細: .github/skills/pptx-slide/references/bullet-fade-animation.md を参照

#:property TargetFramework=net10.0-windows
#:property BuiltInComInteropSupport=true

using System.Reflection;
using System.Runtime.InteropServices;

if (args.Length < 3)
{
    Console.Error.WriteLine("Usage: dotnet run add-bullet-fade-animation.cs <input.pptx> <output.pptx> <slide-index> ...");
    return 1;
}

Console.OutputEncoding = System.Text.Encoding.UTF8;

var inputPath = Path.GetFullPath(args[0]);
var outputPath = Path.GetFullPath(args[1]);
if (!File.Exists(inputPath))
{
    Console.Error.WriteLine($"Error: File not found: {inputPath}");
    return 1;
}

int[] slideIndices;
try
{
    slideIndices = args.Skip(2).Select(int.Parse).ToArray();
}
catch (FormatException)
{
    Console.Error.WriteLine("Error: slide indices must be integers");
    return 1;
}

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine("Error: This script requires Windows + PowerPoint.");
    return 1;
}

const int msoTrue = -1;
const int msoFalse = 0;
const int msoAnimEffectFade = 10;
const int msoAnimateTextByFirstLevel = 2;
const int msoAnimTriggerOnPageClick = 1;
const int ppSaveAsOpenXMLPresentation = 24;

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
    Invoke(app, "Visible", BindingFlags.SetProperty, msoTrue);

    var presentations = Invoke(app, "Presentations", BindingFlags.GetProperty)!;
    presObj = Invoke(presentations, "Open", BindingFlags.InvokeMethod,
        inputPath, msoFalse /*ReadOnly*/, msoFalse /*Untitled*/, msoTrue /*WithWindow*/)!;

    var slides = Invoke(presObj, "Slides", BindingFlags.GetProperty)!;

    foreach (var sIdx in slideIndices)
    {
        var slide = Invoke(slides, "Item", BindingFlags.InvokeMethod, sIdx)!;
        var shapes = Invoke(slide, "Shapes", BindingFlags.GetProperty)!;
        int shapeCount = (int)Invoke(shapes, "Count", BindingFlags.GetProperty)!;
        Console.WriteLine($"Slide {sIdx}: {shapeCount} shapes");

        var timeLine = Invoke(slide, "TimeLine", BindingFlags.GetProperty)!;
        var mainSeq = Invoke(timeLine, "MainSequence", BindingFlags.GetProperty)!;

        int existing = (int)Invoke(mainSeq, "Count", BindingFlags.GetProperty)!;
        for (int e = existing; e >= 1; e--)
        {
            var ex = Invoke(mainSeq, "Item", BindingFlags.InvokeMethod, e)!;
            Invoke(ex, "Delete", BindingFlags.InvokeMethod);
            Marshal.FinalReleaseComObject(ex);
        }
        if (existing > 0) Console.WriteLine($"  Cleared {existing} existing effect(s)");

        int added = 0;
        for (int i = 1; i <= shapeCount; i++)
        {
            var shape = Invoke(shapes, "Item", BindingFlags.InvokeMethod, i)!;
            var name = (string)Invoke(shape, "Name", BindingFlags.GetProperty)!;

            if (name.StartsWith("Content Placeholder"))
            {
                var effect = Invoke(mainSeq, "AddEffect", BindingFlags.InvokeMethod,
                    shape, msoAnimEffectFade, msoAnimateTextByFirstLevel, msoAnimTriggerOnPageClick)!;
                Console.WriteLine($"  Shape {i} ({name}): Fade animation added");
                Marshal.FinalReleaseComObject(effect);
                added++;
            }
            Marshal.FinalReleaseComObject(shape);
        }

        if (added == 0)
        {
            Console.WriteLine($"  (no \"Content Placeholder*\" shape found on slide {sIdx})");
        }

        Marshal.FinalReleaseComObject(mainSeq);
        Marshal.FinalReleaseComObject(timeLine);
        Marshal.FinalReleaseComObject(shapes);
        Marshal.FinalReleaseComObject(slide);
    }

    Invoke(presObj, "SaveAs", BindingFlags.InvokeMethod, outputPath, ppSaveAsOpenXMLPresentation, msoFalse);
    Console.WriteLine($"✅ Saved to {outputPath}");
    return 0;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Error: {ex.Message}");
    Console.Error.WriteLine(ex.StackTrace);
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
