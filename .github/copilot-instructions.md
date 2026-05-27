# Copilot Instructions

## スクリプティングは必ず C# (.NET) を使うこと【最重要】

このリポジトリでは、スクリプト実行は **必ず C# (.NET file-based apps)** で行います。
**Python、PowerShell スクリプト、Bash スクリプト、Node.js などの他のスクリプト言語を使用することは禁止です。**

データ処理・自動化・ワンライナー・調査用スクリプト等、用途を問わず C# を使用してください。

### 必ず守るルール

- ✅ **必ず `dotnet run -` にコードをパイプして実行する**（ファイルを作らない）
- ✅ JSON / CSV / Excel などのデータ処理も C# で行う（`System.Text.Json` / `CsvHelper` / `ClosedXML` 等）
- ❌ `python`, `python3`, `py` コマンドの使用は禁止
- ❌ `.py` ファイルの作成・実行は禁止
- ❌ Bash スクリプト (`.sh`) や Node.js スクリプト (`.js` を `node` で実行) も禁止
- ❌ 一時的な `.cs` ファイルを `Set-Content` / `Out-File` 等で作成してから実行するのも禁止（パイプ実行を使うこと）

### 使い方の概要（PowerShell）

```powershell
'Console.WriteLine("Hello!");' | dotnet run -
```

複数行・外部パッケージ利用・既存ファイル実行などの詳細な書き方は、必ず以下のスキルを参照してください。

👉 **[.github/skills/scripting-guide/SKILL.md](./skills/scripting-guide/SKILL.md)**

スクリプトを書く前に、まず `scripting-guide` スキルの内容を読んで従ってください。

### 例外

ユーザーが明示的に他言語の使用を指示した場合のみ、例外的にその言語を使ってよい。
それ以外のすべてのケースで C# を使うこと。
