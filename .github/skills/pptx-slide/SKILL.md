---
name: pptx-slide
description: 'Microsoft 公式テーマを使った本格的な PowerPoint (.pptx) スライドを作成する。72 種のレイアウトを持つ seed.pptx をベースに、C# スクリプトで unpack → スライド追加 → XML 編集 → pack を行い、Visual QA まで実施。「PowerPoint スライドを作成」「pptx で作って」「テーマ付きスライド」「顧客向けプレゼン資料」「Microsoft テーマ」などの依頼で使用。'
argument-hint: 'スライドの内容（タイトル、構成、保存先）を指定してください'
---

# PPTX テーマ付きスライド作成

## 概要

Microsoft 公式テーマ（`2024-07-29-theme.thmx`）が適用済みの `seed.pptx` をベースに、C# スクリプト（`unpack` / `add-slide` / `pack` / `export-images`）を使って本格的な PowerPoint スライドを作成します。

72 種のスライドレイアウトから適切なものを選び、XML を直接編集してプロフェッショナルなスライドを構築します。Visual QA（サブエージェントによる目視チェック）まで含む完全なワークフローです。

## 前提条件

- **Windows + PowerPoint** がインストールされていること（COM 連携で使用）
- **.NET 10 SDK 以降**（C# file-based apps）

## ファイル構成

スキル assets:

```
.github/skills/pptx-slide/
├── assets/
│   ├── seed.pptx                       # Microsoft テーマ適用済みベース (11.5MB)
│   ├── layouts.md                      # 72 種レイアウトカタログ
│   └── scripts/
│       ├── unpack.cs                   # PPTX 展開 + XML pretty-print
│       ├── pack.cs                     # 展開ディレクトリ → PPTX
│       ├── add-slide.cs                # レイアウト指定スライド追加
│       ├── list-layouts.cs             # レイアウト一覧取得
│       ├── export-images.cs            # JPG エクスポート（QA 用）
│       └── add-bullet-fade-animation.cs # 箇条書きフェードアニメーション追加（オプション）
└── references/
    └── bullet-fade-animation.md        # 箇条書きフェードアニメーションの詳細
```

## ワークフロー

### 1. 保存先決定

- ユーザーに **最終 PPTX の保存先**（ディレクトリとファイル名）を確認する
- 作業中の中間ファイル（unpacked ディレクトリ、JPG など）は `.tmp/` 配下（gitignore 対象）に置く

### 2. seed.pptx を作業ディレクトリにコピー

```powershell
Copy-Item ".github/skills/pptx-slide/assets/seed.pptx" -Destination "<作業ディレクトリ>/<出力ファイル名>.pptx"
```

ここでは、`unpack` 後に編集するため、`.tmp/<作業名>.pptx` にコピーするのが便利です。

### 3. unpack（展開）

```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/unpack.cs <入力.pptx> <展開先ディレクトリ>
```

例:
```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/unpack.cs .tmp/work.pptx .tmp/work-unpacked
```

展開先には `ppt/`, `_rels/`, `[Content_Types].xml` などが含まれます。

### 4. レイアウト選定

スライド構成を計画し、各スライドに使うレイアウト番号を決めます。`assets/layouts.md` の主要レイアウト一覧から選びます。

主要レイアウト（[layouts.md](./assets/layouts.md) 参照）:
- **3**: タイトル スライド（表紙）
- **13**: タイトルとコンテンツ（標準）
- **15**: Two Column Bullet text（2 列箇条書き）
- **22**: タイトルのみ
- **45**: Developer Code Layout（コード）
- **51**: Section Title（セクション区切り）
- **66**: Thank you slide（クロージング）

**⚠️ レイアウトは積極的に変えること**。毎スライド「タイトル + 箇条書き」だけは退屈なスライドの典型。2 列、写真系、コード、引用、セクション区切り等を組み合わせる。

### 5. スライド追加

各スライドを `add-slide.cs` で追加します:

```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/add-slide.cs <展開先ディレクトリ> layout:N
```

例（3 枚追加する場合）:
```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/add-slide.cs .tmp/work-unpacked layout:3   # 表紙
dotnet run .github/skills/pptx-slide/assets/scripts/add-slide.cs .tmp/work-unpacked layout:15  # 2列
dotnet run .github/skills/pptx-slide/assets/scripts/add-slide.cs .tmp/work-unpacked layout:13  # 標準
```

実行すると `ppt/slides/slide1.xml`、`slide2.xml`、`slide3.xml` が生成されます。

### 6. XML 編集（プレースホルダー置換）

各 `ppt/slides/slideN.xml` の placeholder にコンテンツを書き込みます。

**placeholder の種類:**
- `<p:ph type="title"/>` — タイトル
- `<p:ph type="body" idx="N"/>` — 本文
- `<p:ph sz="quarter" idx="N"/>` — 多列レイアウトの各列（`idx` の値で識別）

**編集例（slide1.xml）:**

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="title"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:rPr lang="ja-JP" dirty="0"/><a:t>タイトル文字列</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <!-- 必要な placeholder を追加 -->
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
```

レイアウト本体（`.tmp/work-unpacked/ppt/slideLayouts/slideLayoutN.xml`）を参照すれば、placeholder の `type` / `idx` / 位置サイズが分かります。

### 7. pack（再パック）

XML 編集が終わったら、PPTX に戻します:

```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/pack.cs <展開先ディレクトリ> <出力.pptx>
```

例:
```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/pack.cs .tmp/work-unpacked <保存先>/slides.pptx
```

### 8. Visual QA 用に JPG エクスポート

```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/export-images.cs <出力.pptx> <JPG出力先>
```

例:
```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/export-images.cs <保存先>/slides.pptx .tmp/qa
```

`slide-01.jpg`, `slide-02.jpg`, ... が生成されます。

### 9. サブエージェントによる Visual QA

サブエージェント（general-purpose / explore）に以下のプロンプトで Visual QA を依頼:

```
以下のスライド画像を視覚的に検査してください。問題があることを前提に、バグ探しとして確認してください。

確認ポイント:
- テキストの重なり・はみ出し・切れ
- テキストとボックス境界のオーバーフロー
- 低コントラスト（背景と文字が似た色）
- 要素間の間隔が狭すぎる・広すぎる
- スライドの内容が期待通りか（空欄やプレースホルダーテキストの残留がないか）
- 全体的なレイアウトの問題

画像を読み込んで分析してください:
1. <パス>/slide-01.jpg （期待: ...）
2. <パス>/slide-02.jpg （期待: ...）
3. ...

各スライドの問題点を全て報告してください。
```

### 10. 修正ループ

QA で指摘された問題があれば、XML を編集 → pack → export-images → 再 QA を繰り返します。**最低 1 回は修正→再検証サイクルを回してから完成とすること**。

## オプション: 箇条書きアニメーション

ユーザーから「箇条書きを 1 つずつ表示したい」「アニメーションを付けて」と明示的に依頼された場合のみ、トップレベル箇条書きごとに 1 クリックでフェード表示するアニメーションを追加できます（ぶら下がりの子要素は親と同時表示）。

専用スクリプト `assets/scripts/add-bullet-fade-animation.cs` を使用します。**デフォルトのスライド作成ワークフローには含めません**（必要なときだけ追加します）。

詳細は [references/bullet-fade-animation.md](./references/bullet-fade-animation.md) を参照してください。

## XML 編集のコツ

### 見出しに bullet を付けないようにする

リスト系レイアウト（13, 15 など）では bullet がデフォルト ON です。見出し行（「経歴」「専門分野」など）から bullet を消すには:

```xml
<a:p>
  <a:pPr lvl="0"><a:buNone/></a:pPr>
  <a:r><a:rPr lang="ja-JP" b="1" dirty="0"/><a:t>経歴</a:t></a:r>
</a:p>
```

### 多項目は `<a:p>` を分ける

❌ NG（一つの段落にまとめる）:
```xml
<a:p>
  <a:r><a:t>1. 項目A  2. 項目B  3. 項目C</a:t></a:r>
</a:p>
```

✅ OK（段落を分ける）:
```xml
<a:p><a:pPr lvl="0"/><a:r><a:rPr lang="ja-JP"/><a:t>項目A</a:t></a:r></a:p>
<a:p><a:pPr lvl="0"/><a:r><a:rPr lang="ja-JP"/><a:t>項目B</a:t></a:r></a:p>
<a:p><a:pPr lvl="0"/><a:r><a:rPr lang="ja-JP"/><a:t>項目C</a:t></a:r></a:p>
```

### 日本語テキストの折り返し

日本語は文節区切り判定が弱く、長い文章は不自然に折り返します。対策:
- 文章を短くする（25 文字程度を目安）
- 体言止めにする
- 並列項目は箇条書きに分ける

### HTML 実体参照

`&` は `&amp;` に、`<` は `&lt;` に、`>` は `&gt;` にエスケープしてください。

### bold / italic

- bold: `<a:rPr ... b="1"/>`
- italic: `<a:rPr ... i="1"/>`
- カラー: `<a:rPr ...><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:rPr>`

## 注意事項

- スライドのテキストは **日本語** で作成する
- **絵文字は使用しない**。プロフェッショナルなビジネス資料として、テキストのみで表現する
- 著者名は「Kazuki Ota」、所属は「日本マイクロソフト」（または「Microsoft Japan」）をデフォルトとする
- 中間ファイル（unpacked ディレクトリ、JPG）は `.tmp/` 配下に置き、最終 PPTX のみ作業フォルダーに保存
- 完成前に必ず Visual QA を実施し、最低 1 回は修正サイクルを回す

## トラブルシューティング

### `Presentations.Open : 無効な要求です` エラー

`export-images.cs` で COM 引数の型変換が失敗している。`-1` (msoTrue) などを `dynamic` 経由で渡せない。リフレクション（`InvokeMember`）経由に切り替え済み。最新版を使うこと。

### `Built-in COM has been disabled` エラー

`.NET 10` 以降は COM がデフォルトで無効。`export-images.cs` の先頭で `#:property BuiltInComInteropSupport=true` を設定済み。

### PowerPoint で開けない

通常は `unpack` → `pack` だけでファイルが壊れないはずだが、XML 編集ミスが原因のことが多い。
- `<a:t>` の中身にエスケープ漏れの `&`, `<`, `>` がないか
- 閉じタグの抜け
- placeholder の `type` / `idx` の重複

`pack` 後に PowerShell でファイルを開いてエラーを確認すると判明しやすい:
```powershell
$ppt = New-Object -ComObject PowerPoint.Application
$ppt.Visible = $true
$pres = $ppt.Presentations.Open((Resolve-Path "<出力.pptx>").Path)
```
