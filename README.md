# super-dotnet-new-01

2026.05.29 開催の **超 dotnet new** で使う LT 資料とデモアプリ一式です。

GitHub Copilot CLI / Coding Agent と Aspire を組み合わせて、
「リソースを観測しながらコードのバグにたどり着く」流れを実演します。

## 中身

```
.
├── slides.pptx                LT 用スライド (Microsoft 公式テーマ)
├── demo/                      Aspire ベースの LT デモ (バグ仕込み済み)
│   ├── Demo.AppHost/          AppHost (Redis + API + Web をオーケストレーション)
│   ├── Demo.ApiService/       ASP.NET Core Minimal API
│   ├── Demo.Web/              Blazor Server フロントエンド
│   ├── Demo.ServiceDefaults/  OpenTelemetry / Health Check ほか共通設定
│   └── README.md              デモの動かし方とシナリオ
└── .github/
    ├── copilot-instructions.md  リポジトリ全体の Copilot ルール (C# スクリプト必須)
    └── skills/                  Copilot CLI 用スキル (pptx-slide / scripting-guide)
```

## デモのあらすじ

`demo/` には **わざと壊れた** Aspire アプリが入っています。
すべてのリソースは healthy で起動しますが、Web フロントの一部だけが 404 で赤くなります。

Copilot に「ページの一部だけ赤いエラーになっています。原因を調べて」と頼むと、

1. `aspire describe` でリソース状態を確認
2. `aspire logs apiservice` で 404 のリクエストを発見
3. `Demo.ApiService/Program.cs` のルート定義のタイポを特定

— という、**ソースを総当たりで読まずに実行時の状態から絞り込む** 動きを見せます。

詳しい手順とネタばらしは [`demo/README.md`](./demo/README.md) を参照。

## 動かす

事前準備:

- .NET 10 SDK
- [Aspire CLI](https://learn.microsoft.com/dotnet/aspire/cli/install) (`aspire` コマンド)
- Docker (Redis コンテナ用)

```powershell
cd demo
aspire start
```

ダッシュボード URL が出るのでブラウザで開き、`webfrontend` の `/sessions` を確認してください。

## ライセンス

[MIT License](./LICENSE) — Copyright (c) 2026 Kazuki Ota
