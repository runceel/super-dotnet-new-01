# 箇条書きフェードアニメーション

トップレベル箇条書きごとに 1 クリックでフェード表示し、ぶら下がりの子要素は親と同時に表示する、というアニメーション動作を本文プレースホルダーに追加する手順。

## いつ使うか

- ユーザーが「箇条書きを 1 つずつ出したい」「アニメーションを付けて」と明示的に依頼したとき
- プレゼン中に項目ごとに話を進めたい構成のとき

**デフォルトでは追加しない**。スライド作成のメインワークフローとは独立した「オプション機能」として扱う。

## 動作仕様

- トップレベル箇条書き（`<a:pPr lvl="0"/>` の段落）が **1 クリックずつ** フェードイン
- ぶら下がりの子要素（`lvl="1"` 以上）は親と **同時表示**（クリック不要）
- 箇条書きではない段落（`<a:buNone/>` を持つ見出しや補足文）も **それぞれ別クリック** になる（PowerPoint の自動動作）

技術的には PowerPoint COM の `MainSequence.AddEffect` に以下を指定:
- `EffectId` = `msoAnimEffectFade` (10)
- `Level` = `msoAnimateTextByFirstLevel` (2)
- `Trigger` = `msoAnimTriggerOnPageClick` (1)

これにより XML 上は `presetID="10" presetClass="entr"` の `clickEffect` と `withEffect` が複数生成される。

## 使い方

### 1. PowerPoint を閉じる

スクリプトは COM 経由で PPTX を開いて保存し直すため、対象ファイルが PowerPoint で開かれているとファイルロックで失敗する。必ず PowerPoint を終了してから実行する。

### 2. スクリプト実行

```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/add-bullet-fade-animation.cs <input.pptx> <output.pptx> <slide-index> ...
```

例（slides.pptx のスライド 3, 4, 6 にアニメーションを追加して上書き保存）:

```powershell
# 一旦コピーしてから上書き（入出力が同じパスでも動くが、念のためバックアップ推奨）
Copy-Item slides.pptx .tmp/work.pptx -Force
dotnet run .github/skills/pptx-slide/assets/scripts/add-bullet-fade-animation.cs .tmp/work.pptx slides.pptx 3 4 6
```

### 3. 動作確認

`unpack` してアニメーション XML が入っていることを確認:

```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/unpack.cs slides.pptx .tmp/verify
# slide3.xml に <p:timing> 要素と presetID="10" が含まれることを確認
```

## 対象 Shape

スクリプトは各スライドの **「Content Placeholder」で始まる名前の Shape のみ** を対象にする。

- Title 1 や Date Placeholder などはスキップ
- 自動追加されたスライドの本文プレースホルダーは「Content Placeholder N」という命名になっているので通常はそのまま動く
- カスタム命名や複数の本文ボックスを持つレイアウトでは対象が漏れる可能性がある。その場合はスクリプトの判定ロジック（`name.StartsWith("Content Placeholder")`）を編集する

## 制約事項

- **既存アニメーションは置き換え**：実行時、対象スライドの既存 MainSequence は全てクリアしてから新規追加する。アニメーションをカスタマイズ済みのスライドには使わないこと
- **非箇条書き段落も別クリック**：`<a:buNone/>` で bullet を消した見出し段落や、`endParaRPr` だけの空行も、それぞれ別の click 扱いになる（PowerPoint の `msoAnimateTextByFirstLevel` の仕様）
  - クリック回数を減らしたい場合は、XML 編集で見出し用と本文用を別の Shape に分けるか、スクリプト適用後に PowerPoint の「アニメーションウィンドウ」で手動調整する
- **PowerPoint が必要**：COM 連携のため Windows + PowerPoint 必須

## トラブルシューティング

### ファイルが他のプロセスで使用されています

対象 PPTX が PowerPoint で開かれている。終了してから再実行する:

```powershell
Get-Process -Name "POWERPNT" -ErrorAction SilentlyContinue | Select-Object Id
# 表示された PID を確認して停止
Stop-Process -Id <PID> -Force
```

### 「Content Placeholder*」shape が見つからない

スライドの本文プレースホルダー名が異なる可能性がある。`unpack` してから `slide<N>.xml` の `<p:cNvPr ... name="..."/>` を確認:

```powershell
dotnet run .github/skills/pptx-slide/assets/scripts/unpack.cs <pptx> .tmp/inspect
Select-String -Path .tmp/inspect/ppt/slides/slide<N>.xml -Pattern '<p:cNvPr.*name='
```

該当 Shape 名に合わせてスクリプトの判定を変更する。

### アニメーションを取り消したい

スクリプトをもう一度実行する前に、PowerPoint で対象スライドを開き「アニメーション」タブで全効果を選択して Delete する。または、もう一度スクリプトを走らせれば既存はクリアされる（=同じ Fade で上書きされる）。完全に消すには PowerPoint で手動削除のみ。
