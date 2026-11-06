using System;
using System.Net.WebSockets;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;

internal sealed class AgentSocket : IDisposable
{
    private readonly AutoResetEvent wake = new(false);
    private volatile bool connected;
    public bool Connected => connected;
    public WaitHandle WakeHandle => wake;

    public void Start(Uri apiBase, string token, string agentId, CancellationToken cancel)
    {
        _ = Task.Run(async () =>
        {
            var endpoint = new Uri(new Uri(apiBase.ToString().TrimEnd('/') + "/"),
                "api/print-agent/socket");
            var builder = new UriBuilder(endpoint)
            {
                Scheme = endpoint.Scheme == "https" ? "wss" : "ws",
                Port = endpoint.IsDefaultPort ? -1 : endpoint.Port
            };
            while (!cancel.IsCancellationRequested)
            {
                try
                {
                    using var socket = new ClientWebSocket();
                    socket.Options.SetRequestHeader("Authorization", "Bearer " + token);
                    socket.Options.SetRequestHeader("X-Print-Agent-Id", agentId);
                    socket.Options.KeepAliveInterval = TimeSpan.FromSeconds(20);
                    await socket.ConnectAsync(builder.Uri, cancel);
                    connected = true;
                    wake.Set(); // Read any queued jobs after reconnect.
                    var buffer = new byte[4096];
                    while (!cancel.IsCancellationRequested && socket.State == WebSocketState.Open)
                    {
                        var result = await socket.ReceiveAsync(new ArraySegment<byte>(buffer), cancel);
                        if (result.MessageType == WebSocketMessageType.Close) break;
                        // Notification has no job contents; claim each job from D1 via existing API.
                        if (result.MessageType == WebSocketMessageType.Text) wake.Set();
                        while (!result.EndOfMessage && socket.State == WebSocketState.Open)
                            result = await socket.ReceiveAsync(new ArraySegment<byte>(buffer), cancel);
                    }
                }
                catch (OperationCanceledException) when (cancel.IsCancellationRequested) { break; }
                catch (Exception ex) { Console.Error.WriteLine("WebSocket: " + ex.Message); }
                finally { connected = false; wake.Set(); }
                try { await Task.Delay(TimeSpan.FromSeconds(10), cancel); }
                catch (OperationCanceledException) { break; }
            }
        }, cancel);
    }

    public void Dispose() => wake.Dispose();
}
