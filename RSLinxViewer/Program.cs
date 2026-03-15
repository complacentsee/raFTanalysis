using System.Runtime.InteropServices;
using System.Text;
using Spectre.Console;
using Spectre.Console.Rendering;

namespace RSLinxViewer;

class Program
{
    [DllImport("kernel32.dll")]
    static extern bool SetConsoleOutputCP(uint cp);

    static async Task<int> Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        SetConsoleOutputCP(65001);

        // Parse CLI args: [--driver NAME [IP...]]... [--dll PATH] [--log-dir DIR] [--debug-xml] [--monitor] [--probe]
        var drivers = new List<DriverArg>();
        string? dllPath = null;
        string logDir = @"C:\temp";
        bool debugXml = false;
        bool probeDispids = false;

        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i].ToLowerInvariant();
            if (arg == "--driver" && i + 1 < args.Length)
            {
                i++;
                var drv = new DriverArg { Name = args[i] };
                while (i + 1 < args.Length && !args[i + 1].StartsWith("--"))
                {
                    i++;
                    drv.IPs.Add(args[i]);
                }
                drivers.Add(drv);
            }
            else if (arg == "--dll" && i + 1 < args.Length)
            {
                dllPath = args[++i];
            }
            else if (arg == "--log-dir" && i + 1 < args.Length)
            {
                logDir = args[++i];
            }
            else if (arg == "--debug-xml")
            {
                debugXml = true;
            }
            else if (arg == "--monitor")
            {
                // --monitor is accepted for compatibility but viewer always triggers a full browse
            }
            else if (arg == "--probe")
            {
                probeDispids = true;
            }
        }

        if (drivers.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] RSLinxViewer --driver NAME [[IP...]] [[--driver NAME2 [[IP...]]]] [[--dll PATH]] [[--debug-xml]] [[--probe]]");
            return 1;
        }

        // Resolve DLL path (only needed if injection is required)
        if (dllPath == null)
        {
            string exeDir = AppContext.BaseDirectory;
            dllPath = Path.Combine(exeDir, "RSLinxHook.dll");
        }
        dllPath = Path.GetFullPath(dllPath);

        // Find RSLinx.exe
        AnsiConsole.MarkupLine("[dim]Finding RSLinx.exe...[/]");
        uint pid = Injector.FindRSLinxPid();
        if (pid == 0)
        {
            AnsiConsole.MarkupLine("[red][[FAIL]][/] RSLinx.exe not found. Is it running?");
            return 1;
        }
        AnsiConsole.MarkupLine($"[green][[OK]][/] RSLinx.exe PID: {pid}");

        TryDelete(Path.Combine(logDir, "hook_log.txt"));
        TryDelete(Path.Combine(logDir, "hook_results.txt"));

        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            cts.Cancel();
        };

        using var pipeClient = new PipeClient();

        // Smart injection: probe first, inject only when hook is not already running
        AnsiConsole.MarkupLine("[dim]Probing for existing hook...[/]");
        bool alreadyLoaded = await pipeClient.TryConnectAsync(500, cts.Token);

        if (alreadyLoaded)
        {
            AnsiConsole.MarkupLine("[green][[OK]][/] Hook already running - skipping injection");
        }
        else
        {
            if (!File.Exists(dllPath))
            {
                AnsiConsole.MarkupLine($"[red][[FAIL]][/] RSLinxHook.dll not found at: {Markup.Escape(dllPath)}");
                AnsiConsole.MarkupLine("[dim]Use --dll PATH to specify location[/]");
                return 1;
            }

            AnsiConsole.MarkupLine("[dim]Hook not found, injecting DLL...[/]");
            bool injected = Injector.InjectDLL(pid, dllPath, msg =>
            {
                if (msg.Contains("[FAIL]"))
                    AnsiConsole.MarkupLine($"[red]{Markup.Escape(msg)}[/]");
                else if (msg.Contains("[OK]"))
                    AnsiConsole.MarkupLine($"[green]{Markup.Escape(msg)}[/]");
                else
                    AnsiConsole.MarkupLine($"[dim]{Markup.Escape(msg)}[/]");
            });

            if (!injected)
            {
                AnsiConsole.MarkupLine("[red][[FAIL]][/] DLL injection failed");
                return 1;
            }

            AnsiConsole.MarkupLine("[green][[OK]][/] DLL injected, waiting for pipe server...");
            if (!await pipeClient.TryConnectAsync(10000, cts.Token))
            {
                AnsiConsole.MarkupLine("[red][[FAIL]][/] Hook DLL pipe server did not appear within 10s");
                AnsiConsole.MarkupLine("[dim]Ensure RSLinxHook.dll has pipe support compiled in[/]");
                Cleanup(pipeClient);
                return 1;
            }
        }

        AnsiConsole.MarkupLine("[green][[OK]][/] Connected to hook pipe, sending config...");

        try
        {
            // Always send inject mode — viewer always triggers a full browse cycle
            pipeClient.SendConfig(drivers, logDir, debugXml, monitorMode: false, probeDispids);
        }
        catch (IOException ex)
        {
            AnsiConsole.MarkupLine($"[red][[FAIL]][/] Pipe broken during config send: {Markup.Escape(ex.Message)}");
            Cleanup(pipeClient);
            return 1;
        }
        AnsiConsole.MarkupLine("[green][[OK]][/] Config sent, starting TUI...");
        Thread.Sleep(300);

        var readTask = pipeClient.ReadLoopAsync(cts.Token);

        await RunLiveDisplay(pipeClient, cts.Token);

        Cleanup(pipeClient);
        try { await readTask; } catch { }

        AnsiConsole.MarkupLine("\n[dim]Exited cleanly.[/]");
        return 0;
    }

    static async Task RunLiveDisplay(PipeClient pipe, CancellationToken ct)
    {
        int scrollOffset = 0;

        await AnsiConsole.Live(new Text("Initializing..."))
            .AutoClear(true)
            .StartAsync(async ctx =>
            {
                string? lastXml = null;
                List<string>? lastNodeBlock = null;
                Tree? currentTree = null;

                while (!ct.IsCancellationRequested)
                {
                    while (Console.KeyAvailable)
                    {
                        var key = Console.ReadKey(true);
                        switch (key.Key)
                        {
                            case ConsoleKey.UpArrow:
                                scrollOffset = Math.Max(0, scrollOffset - 1);
                                break;
                            case ConsoleKey.DownArrow:
                                scrollOffset++;
                                break;
                            case ConsoleKey.PageUp:
                                scrollOffset = Math.Max(0, scrollOffset - 10);
                                break;
                            case ConsoleKey.PageDown:
                                scrollOffset += 10;
                                break;
                            case ConsoleKey.Home:
                                scrollOffset = 0;
                                break;
                            case ConsoleKey.B:
                                pipe.SendBrowse();
                                break;
                        }
                    }

                    var nodeBlock = pipe.GetLatestNodeBlock();
                    string? xml = pipe.GetLatestXml();
                    var logLines = pipe.GetRecentLog(15);
                    var (total, identified, events) = pipe.GetStatus();

                    // Prefer N| node block; fall back to X| XML (intermediate polls or debug-xml mode)
                    if (nodeBlock != null && !ReferenceEquals(nodeBlock, lastNodeBlock))
                    {
                        lastNodeBlock = nodeBlock;
                        currentTree = TreeBuilder.Build(nodeBlock);
                    }
                    else if (nodeBlock == null && xml != null && xml != lastXml)
                    {
                        lastXml = xml;
                        currentTree = TopologyParser.BuildTree(xml);
                    }

                    var layout = BuildLayout(currentTree, logLines, total, identified, events, pipe.IsConnected, pipe.IsDone, scrollOffset, out var treeViewport);
                    ctx.UpdateTarget(layout);
                    ctx.Refresh();

                    if (treeViewport != null && treeViewport.TotalLines > 0)
                    {
                        int viewportHeight = Math.Max(5, Console.WindowHeight - Math.Min(logLines.Count, 15) - 8);
                        int maxScroll = Math.Max(0, treeViewport.TotalLines - viewportHeight);
                        scrollOffset = Math.Min(scrollOffset, maxScroll);
                    }

                    try
                    {
                        await Task.Delay(250, ct);
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                }

                // Final render — prefer latest node block, fall back to XML
                var finalNodeBlock = pipe.GetLatestNodeBlock();
                string? finalXml = pipe.GetLatestXml();
                var finalLog = pipe.GetRecentLog(15);
                var finalStatus = pipe.GetStatus();
                Tree? finalTree = finalNodeBlock != null
                    ? TreeBuilder.Build(finalNodeBlock)
                    : (finalXml != null ? TopologyParser.BuildTree(finalXml) : currentTree);
                var finalLayout = BuildLayout(finalTree, finalLog, finalStatus.total, finalStatus.identified, finalStatus.events, false, true, 0, out _);
                ctx.UpdateTarget(finalLayout);
                ctx.Refresh();
            });
    }

    static IRenderable BuildLayout(Tree? tree, List<string> logLines, int total, int identified, int events, bool connected, bool done, int scrollOffset, out Viewport? viewport)
    {
        int consoleHeight = Console.WindowHeight;
        int logLines_ = Math.Min(logLines.Count, 15);
        int treeViewportHeight = Math.Max(5, consoleHeight - logLines_ - 8);

        IRenderable treeContent;
        viewport = null;
        if (tree != null)
        {
            viewport = new Viewport(tree, scrollOffset, treeViewportHeight);
            treeContent = viewport;
        }
        else
        {
            treeContent = new Markup("[dim]Waiting for topology data...[/]");
        }

        string scrollHint = scrollOffset > 0
            ? $" [dim](line {scrollOffset + 1}, \u2191\u2193 PgUp/PgDn)[/]"
            : $" [dim](\u2191\u2193 PgUp/PgDn)[/]";
        var treePanel = new Panel(treeContent)
        {
            Header = new PanelHeader($"[bold]RSLinx Topology[/]{scrollHint}"),
            Border = BoxBorder.Rounded,
            Expand = true,
        };

        var logBuilder = new StringBuilder();
        if (logLines.Count == 0)
        {
            logBuilder.Append("[dim]Waiting for log data...[/]");
        }
        else
        {
            for (int i = 0; i < logLines.Count; i++)
            {
                if (i > 0) logBuilder.AppendLine();
                logBuilder.Append(FormatLogLine(logLines[i]));
            }
        }

        var logPanel = new Panel(new Markup(logBuilder.ToString()))
        {
            Header = new PanelHeader("[bold]Log[/]"),
            Border = BoxBorder.Rounded,
            Expand = true,
        };

        string status = !connected
            ? $"[red]Disconnected[/] | {total} total, {identified} identified, {events} events | [dim]B=re-browse  Ctrl+C=exit[/]"
            : done
                ? $"[dim]Done[/] | {total} total, {identified} identified, {events} events | [dim]B=re-browse  Ctrl+C=exit[/]"
                : $"[green]Browsing[/] | {total} total, {identified} identified, {events} events | [dim]B=re-browse  Ctrl+C=exit[/]";

        return new Rows(treePanel, logPanel, new Markup(status));
    }

    static string FormatLogLine(string line)
    {
        string escaped = Markup.Escape(line);
        if (line.Contains("[FAIL]"))  return $"[red]{escaped}[/]";
        if (line.Contains("[OK]"))    return $"[green]{escaped}[/]";
        if (line.Contains("[WARN]"))  return $"[yellow]{escaped}[/]";
        if (line.Contains("[ENUM:") || line.Contains("[BUS:")) return $"[cyan]{escaped}[/]";
        if (line.Contains("[MONITOR]")) return $"[blue]{escaped}[/]";
        if (line.Contains("[ENGINE]"))  return $"[magenta]{escaped}[/]";
        if (line.StartsWith("==="))     return $"[bold]{escaped}[/]";
        return $"[dim]{escaped}[/]";
    }

    static void Cleanup(PipeClient pipe)
    {
        pipe.SendStop();
        AnsiConsole.MarkupLine("[dim]Waiting for hook to finish...[/]");
        for (int i = 0; i < 20; i++)
        {
            Thread.Sleep(500);
            if (!pipe.IsConnected || pipe.IsDone)
            {
                AnsiConsole.MarkupLine("[dim]Hook finished cleanly.[/]");
                break;
            }
        }
    }

    static void TryDelete(string path)
    {
        try { File.Delete(path); } catch { }
    }
}

class DriverArg
{
    public string Name { get; set; } = "";
    public List<string> IPs { get; set; } = new();
}
