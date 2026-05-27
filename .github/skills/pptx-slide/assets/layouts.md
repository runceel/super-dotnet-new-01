# Microsoft テーマ レイアウトカタログ

`seed.pptx` に含まれる 72 種のスライドレイアウト一覧です。`add-slide.cs` の `layout:N` 引数で指定します。

## よく使うレイアウト（推奨）

| # | 名前 | 用途 |
|---|------|------|
| **3**  | タイトル スライド | プレゼン冒頭の表紙。タイトル + サブタイトル/著者情報 |
| **11** | Blank title | シンプルなタイトルだけのスライド |
| **12** | Agenda | アジェンダ（目次）スライド |
| **13** | タイトルとコンテンツ | 標準的な見出し + 箇条書きスライド |
| **14** | Title & Non-bulleted text | タイトル + 本文（bullet なし） |
| **15** | Two Column Bullet text | 2 列の箇条書きスライド |
| **16** | Two Column Non-bulleted text | 2 列の本文スライド（bullet なし） |
| **22** | タイトルのみ | タイトルのみ表示 |
| **23** | Title Only - left side | 左寄せタイトルのみ |
| **24** | Small title - half page | 小さめタイトル + 半ページ領域 |
| **43** | Title and text side by side | タイトルと本文を横並び |
| **45** | Developer Code Layout | コード表示用 |
| **50** | Demo slide | デモ画面用 |
| **51** | Section Title | セクション区切り |
| **61** | Quote 1 | 引用スライド |
| **66** | Thank you slide | クロージング「Thank you」 |
| **68** | 白紙 | 完全な白紙スライド |
| **70** | Closing logo slide | ロゴのみのクロージング |

## 写真・画像系レイアウト

| # | 名前 |
|---|------|
| 25 | Title - Square Photo |
| 26 | Square Photo |
| 27 | Square Photo 2 |
| 28 | Photo full bleed lower title |
| 29 | Photo full bleed left title |
| 30 | Photo full bleed right title |
| 31 | Top horizontal photo and title |
| 32 | Bottom horizontal photo and title |
| 33 | Two picture content |
| 34 | Three picture content |
| 35 | Four picture content |
| 36 | Two filmstrip photos |
| 37 | Three filmstrip photos |
| 38 | Four filmstrip photos |
| 39 | Five filmstrip photos |
| 40 | Three round photos |
| 41 | Four round photos |
| 42 | Five round photos |
| 62 | Screenshot 1 |
| 63 | Screenshot 2 |
| 64 | Photoslide 1 |
| 65 | Photo slide 2 |

## 多列レイアウト（3〜5 列）

| # | 名前 |
|---|------|
| 17 | Two Column Bullet with Subheads |
| 18 | Two Column Content with Subheads |
| 19 | Three Column Bullet with Subtitles |
| 20 | Four Column Bullet with Subtitles |
| 21 | Five Column Bullet with Subtitles |

## コード系レイアウト

| # | 名前 |
|---|------|
| 45 | Developer Code Layout |
| 46 | Code Bottom |
| 47 | Code Top |
| 48 | Code Right side |
| 49 | Code Left sde |

## セクション区切り系（バリエーション）

| # | 名前 |
|---|------|
| 51 | Section Title |
| 52 | 1_Section Title |
| 53 | 2_Section Title |
| 54 | 3_Section Title |
| 55 | Section Title 2 |
| 56 | 1_Section Title 2 |
| 57 | Section Title 3 |
| 58 | 1_Section Title 3 |
| 59 | 2_Section Title 3 |
| 60 | 3_Section Title 3 |

## タイトル系（バリエーション）

| # | 名前 |
|---|------|
| 1  | Walkin |
| 2  | 1_Walkin |
| 4  | 2_Title Slide |
| 5  | 1_Title Slide |
| 6  | 3_Title Slide |
| 7  | Title Slide 2 |
| 8  | 1_Title Slide 2 |
| 9  | Title Slide 3 |
| 10 | 1_Title Slide 3 |
| 72 | 1_タイトル スライド |

## その他

| # | 名前 |
|---|------|
| 44 | Title and text side by side 2 |
| 67 | 1_Thank you slide |
| 69 | Blank 2 |
| 71 | Black Notes slide Layout |

---

## 使い方

最新の一覧を取得する場合:
```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/list-layouts.cs .github/skills/pptx-slide/assets/seed.pptx
```

スライド追加例:
```powershell
# タイトルスライドを追加
dotnet run .github/skills/pptx-slide/assets/scripts/add-slide.cs <unpacked-dir> layout:3

# 2 列のコンテンツスライドを追加
dotnet run .github/skills/pptx-slide/assets/scripts/add-slide.cs <unpacked-dir> layout:15
```
