namespace Demo.Web;

public class SessionApiClient(HttpClient httpClient)
{
    public Task<SessionEntry?> GetTodaysSpotlightAsync(CancellationToken cancellationToken = default)
        => httpClient.GetFromJsonAsync<SessionEntry>("/sessions/today", cancellationToken);

    public Task<SessionEntry[]?> GetSessionsAsync(CancellationToken cancellationToken = default)
        => httpClient.GetFromJsonAsync<SessionEntry[]>("/sessions", cancellationToken);
}

public record SessionEntry(string Speaker, string Title, string SpeakerUrl);
