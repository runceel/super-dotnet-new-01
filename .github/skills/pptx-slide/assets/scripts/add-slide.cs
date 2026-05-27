// add-slide.cs - 既存スライドまたはスライドレイアウトをコピーして新しいスライドを追加する
// 使用例（スライド複製）:  dotnet run .github/skills/pptx-slide/assets/scripts/add-slide.cs <unpacked-dir> slide:1
// 使用例（レイアウトから）: dotnet run .github/skills/pptx-slide/assets/scripts/add-slide.cs <unpacked-dir> layout:13

using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.IO.Compression;

if (args.Length < 2)
{
    Console.Error.WriteLine("Usage: dotnet run add-slide.cs <unpacked-dir> <slide:N|layout:N>");
    Console.Error.WriteLine("  slide:N   - Duplicate slide number N (1-based)");
    Console.Error.WriteLine("  layout:N  - Add new slide from layout number N");
    return 1;
}

var unpackedDir = Path.GetFullPath(args[0]);
var target      = args[1];

var presentationPath = Path.Combine(unpackedDir, "ppt", "presentation.xml");
var presRelsPath     = Path.Combine(unpackedDir, "ppt", "_rels", "presentation.xml.rels");
var slidesDir        = Path.Combine(unpackedDir, "ppt", "slides");
var slidesRelsDir    = Path.Combine(unpackedDir, "ppt", "slides", "_rels");
var layoutsDir       = Path.Combine(unpackedDir, "ppt", "slideLayouts");
var contentTypesPath = Path.Combine(unpackedDir, "[Content_Types].xml");

if (!File.Exists(presentationPath))
{
    Console.Error.WriteLine($"Error: presentation.xml not found in {unpackedDir}");
    return 1;
}

// ──────────────────────────────────────────────────
// 1. 既存スライドの最大番号を取得
// ──────────────────────────────────────────────────
if (!Directory.Exists(slidesDir)) Directory.CreateDirectory(slidesDir);
if (!Directory.Exists(slidesRelsDir)) Directory.CreateDirectory(slidesRelsDir);

var existingSlides = Directory.GetFiles(slidesDir, "slide*.xml")
    .Where(f => Regex.IsMatch(Path.GetFileName(f), @"^slide\d+\.xml$"))
    .Select(f => int.Parse(Regex.Match(Path.GetFileName(f), @"\d+").Value))
    .OrderBy(n => n)
    .ToList();

int newSlideNum = (existingSlides.Count > 0 ? existingSlides.Max() : 0) + 1;
var newSlideFile    = $"slide{newSlideNum}.xml";
var newSlideRelFile = $"slide{newSlideNum}.xml.rels";
var newSlidePath    = Path.Combine(slidesDir, newSlideFile);
var newSlideRelPath = Path.Combine(slidesRelsDir, newSlideRelFile);

// ──────────────────────────────────────────────────
// 2. ソースを決定してスライド XML をコピー
// ──────────────────────────────────────────────────
string sourceSlideXml;
string sourceSlideRelsXml;
string layoutRelTarget; // rels 内の slideLayout パス

if (target.StartsWith("slide:"))
{
    int srcNum = int.Parse(target.Split(':')[1]);
    var srcPath    = Path.Combine(slidesDir, $"slide{srcNum}.xml");
    var srcRelPath = Path.Combine(slidesRelsDir, $"slide{srcNum}.xml.rels");

    if (!File.Exists(srcPath)) { Console.Error.WriteLine($"Error: slide{srcNum}.xml not found."); return 1; }

    sourceSlideXml     = File.ReadAllText(srcPath);
    sourceSlideRelsXml = File.Exists(srcRelPath) ? File.ReadAllText(srcRelPath) : DefaultSlideRels();
    layoutRelTarget    = ExtractLayoutRelTarget(sourceSlideRelsXml);
}
else if (target.StartsWith("layout:"))
{
    int layoutNum = int.Parse(target.Split(':')[1]);
    var layoutPath = Path.Combine(layoutsDir, $"slideLayout{layoutNum}.xml");
    if (!File.Exists(layoutPath)) { Console.Error.WriteLine($"Error: slideLayout{layoutNum}.xml not found."); return 1; }

    layoutRelTarget    = $"../slideLayouts/slideLayout{layoutNum}.xml";
    sourceSlideXml     = BlankSlideXml();
    sourceSlideRelsXml = MakeSlideRels(layoutRelTarget);
}
else
{
    Console.Error.WriteLine("Error: target must be 'slide:N' or 'layout:N'");
    return 1;
}

// ──────────────────────────────────────────────────
// 3. 新スライドファイルを書き出す
// ──────────────────────────────────────────────────
File.WriteAllText(newSlidePath, sourceSlideXml, new UTF8Encoding(false));
File.WriteAllText(newSlideRelPath, sourceSlideRelsXml, new UTF8Encoding(false));

// ──────────────────────────────────────────────────
// 4. presentation.xml の sldIdLst に追加
// ──────────────────────────────────────────────────
var presDoc = LoadXml(presentationPath);
var nsManager = new XmlNamespaceManager(presDoc.NameTable);
nsManager.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
nsManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

var sldIdLst = presDoc.SelectSingleNode("//p:sldIdLst", nsManager);
if (sldIdLst == null)
{
    // sldIdLst が存在しない場合（スライド0枚のファイル）は sldMasterIdLst の直後に作成する
    var presRoot = presDoc.SelectSingleNode("//p:presentation", nsManager)
        ?? throw new Exception("p:presentation not found in presentation.xml");
    var sldMasterIdLst = presDoc.SelectSingleNode("//p:sldMasterIdLst", nsManager);
    sldIdLst = presDoc.CreateElement("p:sldIdLst",
        "http://schemas.openxmlformats.org/presentationml/2006/main");
    if (sldMasterIdLst != null)
        presRoot.InsertAfter(sldIdLst, sldMasterIdLst);
    else
        presRoot.PrependChild(sldIdLst);
}

// 最大 id を取得
int maxId = sldIdLst.ChildNodes.Cast<XmlNode>()
    .Select(n => int.TryParse(n.Attributes?["id"]?.Value, out var v) ? v : 0)
    .DefaultIfEmpty(256)
    .Max();

// presentation.xml.rels に新スライドのリレーションを追加
var presRelsDoc = LoadXml(presRelsPath);
var relsNs = new XmlNamespaceManager(presRelsDoc.NameTable);
relsNs.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

var relationships = presRelsDoc.SelectSingleNode("/Relationships", relsNs)
    ?? presRelsDoc.SelectSingleNode("//*[local-name()='Relationships']")
    ?? throw new Exception("Relationships not found in presentation.xml.rels");

// 新しいリレーション ID を決定
var existingRIds = presRelsDoc.SelectNodes("//*[@Id]", relsNs)!
    .Cast<XmlNode>()
    .Select(n => n.Attributes!["Id"]!.Value)
    .Where(id => Regex.IsMatch(id, @"^rId\d+$"))
    .Select(id => int.Parse(id[3..]))
    .DefaultIfEmpty(0)
    .Max();
var newRId = $"rId{existingRIds + 1}";

var relElem = presRelsDoc.CreateElement("Relationship",
    "http://schemas.openxmlformats.org/package/2006/relationships");
relElem.SetAttribute("Id", newRId);
relElem.SetAttribute("Type",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide");
relElem.SetAttribute("Target", $"slides/{newSlideFile}");
relationships.AppendChild(relElem);
SaveXml(presRelsDoc, presRelsPath);

// sldId 要素を追加
var sldIdElem = presDoc.CreateElement("p:sldId",
    "http://schemas.openxmlformats.org/presentationml/2006/main");
sldIdElem.SetAttribute("id", (maxId + 1).ToString());
sldIdElem.SetAttribute("id",
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships", newRId);
sldIdLst.AppendChild(sldIdElem);
SaveXml(presDoc, presentationPath);

// ──────────────────────────────────────────────────
// 5. [Content_Types].xml に追加
// ──────────────────────────────────────────────────
var ctDoc = LoadXml(contentTypesPath);
var ctNs  = new XmlNamespaceManager(ctDoc.NameTable);
ctNs.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types");

var types = ctDoc.SelectSingleNode("/ct:Types", ctNs)
    ?? ctDoc.SelectSingleNode("//*[local-name()='Types']")
    ?? throw new Exception("Types not found in [Content_Types].xml");

var override_ = ctDoc.CreateElement("Override",
    "http://schemas.openxmlformats.org/package/2006/content-types");
override_.SetAttribute("PartName", $"/ppt/slides/{newSlideFile}");
override_.SetAttribute("ContentType",
    "application/vnd.openxmlformats-officedocument.presentationml.slide+xml");
types.AppendChild(override_);
SaveXml(ctDoc, contentTypesPath);

Console.WriteLine($"✅ Added slide{newSlideNum} (from {target}) to {unpackedDir}");
Console.WriteLine($"   Add to sldIdLst already done. Total slides: {existingSlides.Count + 1}");
return 0;

// ──────────────────────────────────────────────────
// ヘルパー
// ──────────────────────────────────────────────────
static XmlDocument LoadXml(string path)
{
    var doc = new XmlDocument { PreserveWhitespace = false };
    doc.Load(path);
    return doc;
}

static void SaveXml(XmlDocument doc, string path)
{
    var settings = new XmlWriterSettings
    {
        Indent = true, IndentChars = "  ", NewLineChars = "\n",
        Encoding = new UTF8Encoding(false)
    };
    using var writer = XmlWriter.Create(path, settings);
    doc.Save(writer);
}

static string ExtractLayoutRelTarget(string relsXml)
{
    var m = Regex.Match(relsXml, @"Type=""[^""]*slide[Ll]ayout[^""]*""\s+Target=""([^""]+)""");
    return m.Success ? m.Groups[1].Value : "../slideLayouts/slideLayout1.xml";
}

static string BlankSlideXml() => """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr>
      <p:cNvPr id="1" name=""/>
      <p:cNvGrpSpPr/><p:nvPr/>
    </p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
        <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
    </p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
""";

static string DefaultSlideRels() => """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>
""";

static string MakeSlideRels(string layoutTarget) => $"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    Target="{layoutTarget}"/>
</Relationships>
""";
