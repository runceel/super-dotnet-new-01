# 超dotnet new セッション一覧デモ

`aspire agent init` で GitHub Copilot に Aspire の作法と観測手段を渡したうえで、
**仕込み済みのアプリケーションコードのバグ** を Copilot に発見してもらう LT デモです。

## 構成

```
demo/
├── Demo.AppHost/         Aspire AppHost (オーケストレーター)
├── Demo.ApiService/      ASP.NET Core Minimal API (セッション一覧 API)
├── Demo.Web/             Blazor Server (フロントエンド)
├── Demo.ServiceDefaults/ 共通設定 (OpenTelemetry / Health Check ほか)
└── (Redis "cache" は AppHost が起動時にコンテナで立ち上げる)
```

- `.NET 10` + Aspire AppHost SDK 13.x
- AppHost が Redis (`cache`)、API、Web をまとめてオーケストレーション
- API は `IConnectionMultiplexer` で Redis に書き込み/読み出し
- Web は API を HTTP で叩き、結果をカードで表示

## 仕込みエラー

`Demo.ApiService/Program.cs` の **ルート文字列に typo** を入れてあります。

```csharp
// Demo.ApiService/Program.cs (PLANTED ERROR)
app.MapGet("/session", ...)        // ← 末尾の "s" が抜けている (正しくは "/sessions")
    .WithName("GetSessions");      // ← シンボル名は正しい
```

Web 側 (`Demo.Web/SessionApiClient.cs`) は **正しく** `/sessions` を叩くので:

- すべてのリソースは正常起動 (Redis OK、API OK、Web OK)
- ブラウザで `/sessions` を開ける
- `/sessions/today` は正常 → 「今日のスポットライト」カードは表示される
- `/sessions` は **404** → 「すべてのセッション」セクションだけ赤いエラーバナー

## 動かす

```powershell
cd demo
aspire start
```

`aspire start` を実行すると、ターミナルにダッシュボード URL が表示されます。ブラウザで開き、`webfrontend` の URL から `/sessions` を開いて部分的な失敗を観測してください。

CLI で観測:

```powershell
aspire describe                 # 全リソース healthy 確認
aspire logs apiservice          # GET /sessions が 404 で返るのを発見
```

## Copilot に調査させる

`.agents/skills/aspire/` と `.github/skills/aspire/` に Aspire スキルがインストール済みです。Copilot に以下のように依頼します。

> セッション一覧ページの「すべてのセッション」セクションが赤いエラーになっています。
> Aspire の状態とログを見て、原因を調査してください。

Copilot の期待される動き (Aspire スキルに従う):

1. `aspire describe` でリソース構成と状態を確認 → 全部 healthy
2. `aspire logs apiservice` で API のログを確認 → `GET /sessions` が 404 で返っているのを発見
3. `Demo.ApiService/Program.cs` のルート定義を読み、`app.MapGet("/session", ...)` (末尾 s 抜け) を発見
4. 修正案を提示

**ポイント**: Copilot は最初にソースを総当たりで読むのではなく、まず `aspire describe` と `aspire logs` で **実行時の状態** を観測してからソースを絞り込みにいく動きが期待値。Aspire スキルがこの順番を教えています。

## 修正

`Demo.ApiService/Program.cs` の一文字を直すだけです。

```diff
-app.MapGet("/session", async (IConnectionMultiplexer redis, CancellationToken ct) =>
+app.MapGet("/sessions", async (IConnectionMultiplexer redis, CancellationToken ct) =>
```

再度 `aspire start` で動作確認。`/sessions` ページに「今日のスポットライト」とすべてのセッション一覧が表示されます。
