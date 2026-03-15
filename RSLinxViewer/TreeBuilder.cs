using Spectre.Console;

namespace RSLinxViewer;

/// <summary>
/// Builds a Spectre.Console Tree from N| topology pipe messages.
///
/// Wire format (lines between N|BEGIN and N|END):
///   N|ROOT|name|classname
///   N|BUS|name|classname
///   N|ADDR|String|ip|deviceName|classname
///   N|ADDR|Short|slot|deviceName|classname
///   N|PUSH    — enter sub-tree (last device has backplane)
///   N|POP     — return to parent bus
/// </summary>
static class TreeBuilder
{
    public static Tree? Build(IReadOnlyList<string>? lines)
    {
        if (lines == null || lines.Count == 0)
            return null;

        Tree? rootTree = null;
        // Stack tracks bus/device nodes so PUSH/POP can navigate depth.
        // Each frame: (containerNode, pendingSlotLeaves)
        var stack = new Stack<(IHasTreeNodes container, List<(string slot, string name, string cls)> leaves)>();

        IHasTreeNodes? current = null;
        var slotLeaves = new List<(string slot, string name, string cls)>();
        TreeNode? lastDevNode = null;  // last device node added (for PUSH targeting)

        foreach (var raw in lines)
        {
            var line = raw.TrimEnd('\n', '\r');
            if (line.Length == 0) continue;

            var parts = line.Split('|');
            if (parts.Length < 2 || parts[0] != "N") continue;

            switch (parts[1])
            {
                case "ROOT" when parts.Length >= 4:
                {
                    string name = parts[2];
                    string cls  = parts[3];
                    rootTree = new Tree(FormatRoot(name, cls));
                    current = rootTree;
                    lastDevNode = null;
                    break;
                }

                case "BUS" when parts.Length >= 4:
                {
                    if (current == null) break;
                    // Flush any queued slot leaves to the current bus before adding a new bus
                    FlushSlotLeaves(slotLeaves, current);
                    string name = parts[2];
                    string cls  = parts[3];
                    // When stack is empty we're at driver level — always attach to root,
                    // not to the previous driver's bus node (which is what current points
                    // to after a POP that drains the stack).
                    IHasTreeNodes parent = stack.Count == 0 ? (IHasTreeNodes)rootTree! : current;
                    var busNode = parent.AddNode(FormatBus(name, cls));
                    current = busNode;
                    lastDevNode = null;
                    break;
                }

                case "ADDR" when parts.Length >= 6 && parts[2] == "String":
                {
                    if (current == null) break;
                    string ip      = parts[3];
                    string devName = parts[4];
                    string cls     = parts[5];
                    lastDevNode = current.AddNode(FormatEthernet(ip, devName, cls));
                    break;
                }

                case "ADDR" when parts.Length >= 6 && parts[2] == "Short":
                {
                    if (current == null) break;
                    string slot    = parts[3];
                    string devName = parts[4];
                    string cls     = parts[5];
                    slotLeaves.Add((slot, devName, cls));
                    lastDevNode = null;  // short addrs don't serve as PUSH targets
                    break;
                }

                case "PUSH":
                {
                    if (current == null) break;
                    // The device that will own the sub-tree is lastDevNode.
                    // Push current bus + its pending slot leaves onto the stack,
                    // then set current = that device node.
                    FlushSlotLeaves(slotLeaves, current);
                    stack.Push((current, new List<(string, string, string)>(slotLeaves)));
                    slotLeaves.Clear();
                    // Navigate into the last Ethernet device node
                    if (lastDevNode != null)
                        current = lastDevNode;
                    lastDevNode = null;
                    break;
                }

                case "POP":
                {
                    if (stack.Count == 0) break;
                    if (current != null) FlushSlotLeaves(slotLeaves, current);
                    slotLeaves.Clear();
                    var (parentContainer, parentLeaves) = stack.Pop();
                    current = parentContainer;
                    slotLeaves = parentLeaves;
                    lastDevNode = null;
                    break;
                }
            }
        }

        // Flush any remaining slot leaves
        if (current != null && slotLeaves.Count > 0)
            FlushSlotLeaves(slotLeaves, current);

        return rootTree;
    }

    // Render accumulated slot leaves as a compact 2-column grid under the bus node.
    static void FlushSlotLeaves(List<(string slot, string name, string cls)> leaves, IHasTreeNodes parent)
    {
        if (leaves.Count == 0) return;

        var grid = new Grid();
        grid.AddColumn(new GridColumn().NoWrap().PadRight(3));
        grid.AddColumn(new GridColumn().NoWrap());

        for (int i = 0; i < leaves.Count; i += 2)
        {
            var (s1, n1, c1) = leaves[i];
            var col1 = new Markup($"[blue][[{Markup.Escape(s1)}]][/] {FormatDeviceName(n1, c1)}");

            if (i + 1 < leaves.Count)
            {
                var (s2, n2, c2) = leaves[i + 1];
                grid.AddRow(col1, new Markup($"[blue][[{Markup.Escape(s2)}]][/] {FormatDeviceName(n2, c2)}"));
            }
            else
            {
                grid.AddRow(col1, new Text(""));
            }
        }

        parent.AddNode(grid);
        leaves.Clear();
    }

    static string FormatRoot(string name, string cls) =>
        $"[bold]{Markup.Escape(name)}[/] [dim]({Markup.Escape(cls)})[/]";

    static string FormatBus(string name, string cls) =>
        string.IsNullOrEmpty(cls)
            ? $"[yellow]{Markup.Escape(name)}[/]"
            : $"[yellow]{Markup.Escape(name)}[/] [dim]({Markup.Escape(cls)})[/]";

    static string FormatEthernet(string ip, string devName, string cls) =>
        $"[cyan]{Markup.Escape(ip)}[/] - {FormatDeviceName(devName, cls)}";

    static string FormatDeviceName(string name, string cls)
    {
        if (cls == "Unrecognized Device")
            return $"[dim italic]{Markup.Escape(name)}[/]";
        return $"[white]{Markup.Escape(name)}[/]";
    }
}
