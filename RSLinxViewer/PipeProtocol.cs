using System.IO.Pipes;
using System.Text;

namespace RSLinxViewer;

/// <summary>
/// Bidirectional named pipe client for RSLinxHook.dll communication.
/// Protocol: UTF-8, line-based with type prefix:
///
/// Client → Hook (after connection):
///   C|KEY=VALUE  — config line
///   C|END        — config complete, hook proceeds
///   STOP         — stop signal (on Ctrl+C)
///
/// Hook → Client:
///   L|text       — log line
///   S|t|i|e      — status (total|identified|events)
///   X|BEGIN       — start of topology XML block
///   ...xml...    — raw XML lines (inside block)
///   X|END        — end of topology XML block
///   D|           — done signal
/// </summary>
sealed class PipeClient : IDisposable
{
    const string PipeName = "RSLinxHook";
    const int MaxLogLines = 200;

    readonly NamedPipeClientStream _pipe;
    readonly object _lock = new();

    readonly List<string> _logLines = new(MaxLogLines);
    string? _latestXml;
    List<string>? _latestNodeBlock;
    int _totalDevices;
    int _identifiedDevices;
    int _eventCount;
    bool _done;
    bool _connected;

    public bool IsConnected => _connected;
    public bool IsDone => _done;

    public PipeClient()
    {
        _pipe = new NamedPipeClientStream(
            ".",
            PipeName,
            PipeDirection.InOut,
            PipeOptions.Asynchronous);
    }

    /// <summary>
    /// Attempt to connect to the hook DLL's pipe server within the given timeout.
    /// Returns true if connected, false on timeout or cancellation.
    /// </summary>
    public async Task<bool> TryConnectAsync(int timeoutMs, CancellationToken ct)
    {
        try
        {
            await _pipe.ConnectAsync(timeoutMs, ct);
            _connected = true;
            return true;
        }
        catch (TimeoutException) { return false; }
        catch (OperationCanceledException) { return false; }
    }

    /// <summary>Send config to hook DLL over pipe. Call after connection established.</summary>
    public void SendConfig(List<DriverArg> drivers, string logDir, bool debugXml, bool monitorMode, bool probeDispids = false)
    {
        var sb = new StringBuilder();
        sb.AppendLine(monitorMode ? "C|MODE=monitor" : "C|MODE=inject");
        if (logDir != @"C:\temp")
            sb.AppendLine($"C|LOGDIR={logDir}");
        if (debugXml)
            sb.AppendLine("C|DEBUGXML=1");
        if (probeDispids)
            sb.AppendLine("C|PROBE=1");
        foreach (var drv in drivers)
        {
            sb.AppendLine($"C|DRIVER={drv.Name}");
            foreach (var ip in drv.IPs)
                sb.AppendLine($"C|IP={ip}");
        }
        sb.AppendLine("C|END");
        byte[] bytes = Encoding.UTF8.GetBytes(sb.ToString());
        _pipe.Write(bytes, 0, bytes.Length);
    }

    /// <summary>Send re-browse command to hook DLL.</summary>
    public void SendBrowse()
    {
        lock (_lock) { _done = false; }
        try
        {
            byte[] bytes = Encoding.UTF8.GetBytes("B|\n");
            _pipe.Write(bytes, 0, bytes.Length);
        }
        catch { /* pipe may already be closed */ }
    }

    /// <summary>Send stop signal to hook DLL.</summary>
    public void SendStop()
    {
        try
        {
            byte[] bytes = Encoding.UTF8.GetBytes("STOP\n");
            _pipe.Write(bytes, 0, bytes.Length);
        }
        catch { /* pipe may already be closed */ }
    }

    /// <summary>
    /// Main read loop — reads lines from pipe and routes by prefix.
    /// Runs until pipe disconnects, done signal, or cancellation.
    /// Runs on a thread-pool thread (sync I/O on non-overlapped handle).
    /// </summary>
    public Task ReadLoopAsync(CancellationToken ct)
    {
        return Task.Run(() =>
        {
            using var reader = new StreamReader(_pipe, Encoding.UTF8, leaveOpen: true);
            var xmlBuilder = new StringBuilder();
            bool inXmlBlock = false;
            var nodeLines = new List<string>();
            bool inNodeBlock = false;

            while (!ct.IsCancellationRequested)
            {
                string? line;
                try
                {
                    line = reader.ReadLine();
                }
                catch (IOException)
                {
                    break; // pipe broken
                }
                catch (ObjectDisposedException)
                {
                    break; // pipe closed
                }

                if (line == null)
                    break; // EOF / disconnect

                if (inXmlBlock)
                {
                    if (line == "X|END")
                    {
                        inXmlBlock = false;
                        lock (_lock)
                        {
                            _latestXml = xmlBuilder.ToString();
                        }
                    }
                    else
                    {
                        xmlBuilder.AppendLine(line);
                    }
                    continue;
                }

                if (inNodeBlock)
                {
                    if (line == "N|END")
                    {
                        inNodeBlock = false;
                        lock (_lock)
                        {
                            _latestNodeBlock = new List<string>(nodeLines);
                        }
                    }
                    else
                    {
                        nodeLines.Add(line);
                    }
                    continue;
                }

                if (line.StartsWith("L|"))
                {
                    string msg = line[2..];
                    lock (_lock)
                    {
                        _logLines.Add(msg);
                        if (_logLines.Count > MaxLogLines)
                            _logLines.RemoveAt(0);
                    }
                }
                else if (line.StartsWith("S|"))
                {
                    var parts = line[2..].Split('|');
                    if (parts.Length >= 3)
                    {
                        lock (_lock)
                        {
                            int.TryParse(parts[0], out _totalDevices);
                            int.TryParse(parts[1], out _identifiedDevices);
                            int.TryParse(parts[2], out _eventCount);
                        }
                    }
                }
                else if (line == "X|BEGIN")
                {
                    inXmlBlock = true;
                    xmlBuilder.Clear();
                }
                else if (line == "N|BEGIN")
                {
                    inNodeBlock = true;
                    nodeLines.Clear();
                }
                else if (line.StartsWith("D|"))
                {
                    lock (_lock) { _done = true; }
                    // session continues — pipe stays open for B|/STOP
                }
            }

            _connected = false;
        }, CancellationToken.None);
    }

    /// <summary>Get the last N log lines (thread-safe copy).</summary>
    public List<string> GetRecentLog(int count)
    {
        lock (_lock)
        {
            int start = Math.Max(0, _logLines.Count - count);
            return _logLines.GetRange(start, _logLines.Count - start);
        }
    }

    /// <summary>Get the latest topology XML string, or null if none received yet.</summary>
    public string? GetLatestXml()
    {
        lock (_lock)
        {
            return _latestXml;
        }
    }

    /// <summary>Get the latest node block lines (N| messages between N|BEGIN and N|END), or null if none received yet.</summary>
    public List<string>? GetLatestNodeBlock()
    {
        lock (_lock)
        {
            return _latestNodeBlock;
        }
    }

    /// <summary>Get current device status counts.</summary>
    public (int total, int identified, int events) GetStatus()
    {
        lock (_lock)
        {
            return (_totalDevices, _identifiedDevices, _eventCount);
        }
    }

    public void Dispose()
    {
        _pipe.Dispose();
    }
}
