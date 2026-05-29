using StackExchange.Redis;
using System.Text.Json;

var builder = WebApplication.CreateBuilder(args);

// Add service defaults & Aspire client integrations.
builder.AddServiceDefaults();

// Add services to the container.
builder.Services.AddProblemDetails();

// Learn more about configuring OpenAPI at https://aka.ms/aspnet/openapi
builder.Services.AddOpenApi();

// Aspire Redis client integration
builder.AddRedisClient("cache");

var app = builder.Build();

// Configure the HTTP request pipeline.
app.UseExceptionHandler();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

SessionEntry[] sessions =
[
    new("@hr_sao", "C++とC#の話（予定）", "https://x.com/hr_sao"),
    new("@okazuki", "Aspire と GitHub Copilot の連携", "https://x.com/okazuki"),
    new("@neuecc", ".NET for Android と Native AOT で Android アプリを Cursor でライブコーディング", "https://x.com/neuecc"),
    new("@tomo_kusaba", "GitHub Copilot CLIでWebアクセシビリティを改善した話", "https://x.com/tomo_kusaba"),
    new("@openlibsys", "ニッチを極めて？オープンソースソフトウェアで起業した話", "https://x.com/openlibsys"),
    new("@mayuki", "AIで雑にWebアプリを作る話", "https://x.com/mayuki"),
    new("@naoki0311", "Copilot Coworkか、その周辺のテクノロジー", "https://x.com/naoki0311"),
    new("@MogamiTsuchikaw", "Avaloniaでデスクトップ開発、もっとみんなやってもいいんじゃない？", "https://x.com/MogamiTsuchikaw"),
    new("@nuskey8", "最近のC#/Unityの開発環境周りの話(予定)", "https://x.com/nuskey8"),
    new("@Akeit0_", "C# JavaScript Engine/Browserの内部の話", "https://x.com/Akeit0_"),
    new("Takao Tetsuro", "Data structuring strategy", "https://mvp.microsoft.com/en-US/mvp/profile/e78d78ba-3c9a-e411-93f2-9cb65495d3c4"),
    new("@muo_jp", "C#とC#のinteropの話", "https://x.com/muo_jp"),
    new("@xin9le", "OpenTelemetry関連の話かInterceptors関連か(予定)", "https://x.com/xin9le"),
    new("@zolic8", "dotnet/runtimeの最近の動向(予定)", "https://x.com/zolic8"),
    new("ycanardeau", "ChatGPTでソースジェネレーターを3日で作った話(予定)", "https://github.com/ycanardeau"),
    new("@ruccho_vector", "Roslynのフロー解析APIについて", "https://x.com/ruccho_vector"),
];

app.MapGet("/", () => "超dotnet new セッション一覧 API. /sessions または /sessions/today を見てください。");

app.MapGet("/sessions/today", async (IConnectionMultiplexer redis, CancellationToken ct) =>
{
    var db = redis.GetDatabase();
    const string key = "session-of-the-day";

    var cached = await db.StringGetAsync(key);
    if (cached.HasValue)
    {
        return JsonSerializer.Deserialize<SessionEntry>(cached.ToString())!;
    }

    var pick = sessions[Random.Shared.Next(sessions.Length)];
    await db.StringSetAsync(key, JsonSerializer.Serialize(pick), TimeSpan.FromMinutes(1));
    return pick;
})
.WithName("GetSessionOfTheDay");

app.MapGet("/session", async (IConnectionMultiplexer redis, CancellationToken ct) =>
{
    var db = redis.GetDatabase();
    const string key = "sessions";

    var cached = await db.StringGetAsync(key);
    if (cached.HasValue)
    {
        return JsonSerializer.Deserialize<SessionEntry[]>(cached.ToString())!;
    }

    await db.StringSetAsync(key, JsonSerializer.Serialize(sessions), TimeSpan.FromHours(1));
    return sessions;
})
.WithName("GetSessions");

app.MapDefaultEndpoints();

app.Run();

record SessionEntry(string Speaker, string Title, string SpeakerUrl);
