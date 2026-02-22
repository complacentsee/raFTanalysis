using System.Text;
using Spectre.Console;
using Spectre.Console.Rendering;

namespace RSLinxViewer;

class Program
{
    static async Task<int> Main(string[] args)
    {
        // Parse CLI args: [--driver NAME [IP...]]... [--dll PATH] [--log-dir DIR] [--debug-xml] [--monitor]
        var drivers = new List<DriverArg>();
        string? dllPath = null;
        string logDir = @"C:\temp";
        bool debugXml = false;
        bool monitorMode = false;
        bool probeDispids = false;

        for (int i = 0; i < args.Length; i++)
        {
            string arg = args[i].ToLowerInvariant();
            if (arg == "--driver" && i + 1 < args.Length)
            {
                i++;
                var drv = new DriverArg { Name = args[i] };
                // Collect IPs until next flag
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
                monitorMode = true;
            }
            else if (arg == "--probe")
            {
                probeDispids = true;
            }
        }

        if (drivers.Count == 0)
        {
            AnsiConsole.MarkupLine("[red]Usage:[/] RSLinxViewer --driver NAME [[IP...]] [[--driver NAME2 [[IP...]]]] [[--dll PATH]] [[--debug-xml]] [[--monitor]] [[--probe]]");
            return 1;
        }

        // Resolve DLL path
        if (dllPath == null)
        {
            // Look next to our exe
            string exeDir = AppContext.BaseDirectory;
            dllPath = Path.Combine(exeDir, "RSLinxHook.dll");
        }
        if (!File.Exists(dllPath))
        {
            AnsiConsole.MarkupLine($"[red][[FAIL]][/] RSLinxHook.dll not found at: {Markup.Escape(dllPath)}");
            AnsiConsole.MarkupLine("[dim]Use --dll PATH to specify location[/]");
            return 1;
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

        // Delete old output files
        TryDelete(Path.Combine(logDir, "hook_log.txt"));
        TryDelete(Path.Combine(logDir, "hook_results.txt"));

        // Create named pipe server BEFORE injection
        using var pipeServer = new PipeServer();
        AnsiConsole.MarkupLine("[dim]Pipe server created, injecting DLL...[/]");

        // Eject old DLL if present
        Injector.EjectDLL(pid, dllPath, msg => AnsiConsole.MarkupLine($"[dim]{Markup.Escape(msg)}[/]"));
        Thread.Sleep(500);

        // Inject
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

        AnsiConsole.MarkupLine("[green][[OK]][/] DLL injected, waiting for pipe connection...");

        // Wait for hook to connect to our pipe
        using var cts = new CancellationTokenSource();
        Console.CancelKeyPress += (_, e) =>
        {
            e.Cancel = true;
            cts.Cancel();
        };

        try
        {
            var connectTimeout = new CancellationTokenSource(TimeSpan.FromSeconds(10));
            var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(cts.Token, connectTimeout.Token);
            await pipeServer.WaitForConnectionAsync(linkedCts.Token);
        }
        catch (OperationCanceledException) when (!cts.IsCancellationRequested)
        {
            AnsiConsole.MarkupLine("[red][[FAIL]][/] Hook DLL did not connect to pipe within 10s");
            AnsiConsole.MarkupLine("[dim]Ensure RSLinxHook.dll has pipe support compiled in[/]");
            Cleanup(pid, dllPath, pipeServer);
            return 1;
        }
        catch (OperationCanceledException)
        {
            Cleanup(pid, dllPath, pipeServer);
            return 0;
        }

        AnsiConsole.MarkupLine("[green][[OK]][/] Hook connected via pipe, sending config...");

        // Send config over pipe (replaces hook_config.txt)
        try
        {
            pipeServer.SendConfig(drivers, logDir, debugXml, monitorMode, probeDispids);
        }
        catch (IOException ex)
        {
            AnsiConsole.MarkupLine($"[red][[FAIL]][/] Pipe broken during config send: {Markup.Escape(ex.Message)}");
            Cleanup(pid, dllPath, pipeServer);
            return 1;
        }
        AnsiConsole.MarkupLine("[green][[OK]][/] Config sent, starting TUI...");
        Thread.Sleep(300); // brief pause so user can see the message

        // Start pipe read loop in background
        var readTask = pipeServer.ReadLoopAsync(cts.Token);

        // Enter TUI live display
        await RunLiveDisplay(pipeServer, cts.Token);

        // Cleanup
        Cleanup(pid, dllPath, pipeServer);

        // Wait for read task
        try { await readTask; } catch { /* ignore */ }

        AnsiConsole.MarkupLine("\n[dim]Exited cleanly.[/]");
        return 0;
    }

    static async Task RunLiveDisplay(PipeServer pipe, CancellationToken ct)
    {
        int scrollOffset = 0;

        await AnsiConsole.Live(new Text("Initializing..."))
            .AutoClear(true)
            .StartAsync(async ctx =>
            {
                string? lastXml = null;
                Tree? currentTree = null;

                while (!ct.IsCancellationRequested && !pipe.IsDone)
                {
                    // Handle keyboard input for scrolling
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
                        }
                    }

                    // Get latest data from pipe
                    string? xml = pipe.GetLatestXml();
                    var logLines = pipe.GetRecentLog(15);
                    var (total, identified, events) = pipe.GetStatus();

                    // Rebuild tree only when XML changes
                    if (xml != null && xml != lastXml)
                    {
                        lastXml = xml;
                        currentTree = TopologyParser.BuildTree(xml);
                    }

                    // Build layout
                    var layout = BuildLayout(currentTree, logLines, total, identified, events, pipe.IsConnected, scrollOffset, out var treeViewport);
                    ctx.UpdateTarget(layout);
                    ctx.Refresh();

                    // Clamp scroll offset to actual tree height
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

                // Final render
                string? finalXml = pipe.GetLatestXml();
                var finalLog = pipe.GetRecentLog(15);
                var finalStatus = pipe.GetStatus();
                Tree? finalTree = finalXml != null ? TopologyParser.BuildTree(finalXml) : currentTree;
                var finalLayout = BuildLayout(finalTree, finalLog, finalStatus.total, finalStatus.identified, finalStatus.events, false, 0, out _);
                ctx.UpdateTarget(finalLayout);
                ctx.Refresh();
            });
    }

    static IRenderable BuildLayout(Tree? tree, List<string> logLines, int total, int identified, int events, bool connected, int scrollOffset, out Viewport? viewport)
    {
        // Calculate viewport height for tree (leave room for log panel + status)
        int consoleHeight = Console.WindowHeight;
        int logLines_ = Math.Min(logLines.Count, 15);
        int treeViewportHeight = Math.Max(5, consoleHeight - logLines_ - 8);

        // Tree panel with viewport scrolling
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

        string scrollHint = scrollOffset > 0 ? $" [dim](line {scrollOffset + 1}, {"\u2191\u2193"} PgUp/PgDn to scroll)[/]" : $" [dim]({"\u2191\u2193"} PgUp/PgDn to scroll)[/]";
        var treePanel = new Panel(treeContent)
        {
            Header = new PanelHeader($"[bold]RSLinx Topology[/]{scrollHint}"),
            Border = BoxBorder.Rounded,
            Expand = true,
        };

        // Log panel
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

        // Status footer
        string status = connected
            ? $"[green]Connected[/] | {total} total, {identified} identified, {events} events | [dim]Ctrl+C to exit[/]"
            : $"[red]Disconnected[/] | {total} total, {identified} identified, {events} events";

        // Combine with rows layout
        var rows = new Rows(
            treePanel,
            logPanel,
            new Markup(status)
        );

        return rows;
    }

    static string FormatLogLine(string line)
    {
        string escaped = Markup.Escape(line);

        // Color-code by tag
        if (line.Contains("[FAIL]"))
            return $"[red]{escaped}[/]";
        if (line.Contains("[OK]"))
            return $"[green]{escaped}[/]";
        if (line.Contains("[WARN]"))
            return $"[yellow]{escaped}[/]";
        if (line.Contains("[ENUM:") || line.Contains("[BUS:"))
            return $"[cyan]{escaped}[/]";
        if (line.Contains("[MONITOR]"))
            return $"[blue]{escaped}[/]";
        if (line.Contains("[ENGINE]"))
            return $"[magenta]{escaped}[/]";
        if (line.StartsWith("==="))
            return $"[bold]{escaped}[/]";

        return $"[dim]{escaped}[/]";
    }

    static void Cleanup(uint pid, string dllPath, PipeServer pipe)
    {
        // Signal hook to stop via pipe
        pipe.SendStop();

        // Wait for hook to finish (pipe disconnect or done signal)
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

        // Brief pause for DLL to fully wind down after cleanup
        Thread.Sleep(500);

        // Eject DLL
        Injector.EjectDLL(pid, dllPath, _ => { });
    }

    static void TryDelete(string path)
    {
        try { File.Delete(path); } catch { /* ignore */ }
    }
}

class DriverArg
{
    public string Name { get; set; } = "";
    public List<string> IPs { get; set; } = new();
}
