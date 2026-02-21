using System.Xml.Linq;
using Spectre.Console;

namespace RSLinxViewer;

/// <summary>
/// Parses Rockwell topology XML (from Harmony.TopologyServices)
/// into a Spectre.Console Tree widget for TUI display.
/// </summary>
static class TopologyParser
{
    /// <summary>
    /// Build a Spectre.Console Tree from topology XML string.
    /// Returns null if XML is invalid or empty.
    /// </summary>
    public static Tree? BuildTree(string? xml)
    {
        if (string.IsNullOrWhiteSpace(xml))
            return null;

        try
        {
            var doc = XDocument.Parse(xml);

            var treeEl = doc.Root?.Element("tree");
            if (treeEl == null)
                return null;

            var rootDevice = treeEl.Element("device");
            if (rootDevice == null)
                return null;

            string rootName = rootDevice.Attribute("name")?.Value ?? "Unknown";
            string rootClass = rootDevice.Attribute("classname")?.Value ?? "";

            var tree = new Tree($"[bold]{Markup.Escape(rootName)}[/] [dim]({Markup.Escape(rootClass)})[/]");

            foreach (var port in rootDevice.Elements("port"))
            {
                ProcessPort(port, tree);
            }

            return tree;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Count devices in topology XML.
    /// </summary>
    public static (int total, int identified) CountDevices(string? xml)
    {
        if (string.IsNullOrWhiteSpace(xml))
            return (0, 0);

        try
        {
            var doc = XDocument.Parse(xml);
            int total = 0, identified = 0;
            foreach (var dev in doc.Descendants("device"))
            {
                // Skip references
                if (dev.Attribute("reference") != null)
                    continue;
                string? cn = dev.Attribute("classname")?.Value;
                if (cn == null)
                    continue;
                total++;
                if (cn != "Unrecognized Device" && cn != "Workstation")
                    identified++;
            }
            return (total, identified);
        }
        catch
        {
            return (0, 0);
        }
    }

    static void ProcessPort(XElement port, IHasTreeNodes parent)
    {
        // A port typically leads to a bus. Show the bus, not the port itself.
        foreach (var bus in port.Elements("bus"))
        {
            ProcessBus(bus, parent);
        }
    }

    static void ProcessBus(XElement bus, IHasTreeNodes parent)
    {
        string busName = bus.Attribute("name")?.Value ?? "Bus";
        string busClass = bus.Attribute("classname")?.Value ?? "";

        // Separate leaf slot modules from other addresses for compact rendering
        var slotLeaves = new List<(string slot, string name, string? classname)>();
        var otherAddresses = new List<XElement>();

        foreach (var addr in bus.Elements("address"))
        {
            string addrType = addr.Attribute("type")?.Value ?? "";
            if (addrType == "Empty") continue;

            var device = addr.Element("device");
            if (device == null || device.Attribute("reference") != null) continue;

            string? devName = device.Attribute("name")?.Value;
            if (devName == null) continue;

            // Leaf slot modules: Short address with no sub-ports
            if (addrType == "Short" && !device.Elements("port").Any())
            {
                slotLeaves.Add((
                    addr.Attribute("value")?.Value ?? "?",
                    devName,
                    device.Attribute("classname")?.Value));
            }
            else
            {
                otherAddresses.Add(addr);
            }
        }

        var busNode = parent.AddNode($"[yellow]{Markup.Escape(busName)}[/] [dim]({Markup.Escape(busClass)})[/]");

        // Non-leaf / non-slot devices render as normal tree nodes
        foreach (var addr in otherAddresses)
            ProcessAddress(addr, busNode);

        // Leaf slot modules render in a compact 2-column grid
        if (slotLeaves.Count > 0)
        {
            var grid = new Grid();
            grid.AddColumn(new GridColumn().NoWrap().PadRight(3));
            grid.AddColumn(new GridColumn().NoWrap());

            for (int i = 0; i < slotLeaves.Count; i += 2)
            {
                var (s1, n1, c1) = slotLeaves[i];
                var col1 = new Markup($"[blue][[{Markup.Escape(s1)}]][/] {FormatDeviceName(n1, c1)}");

                if (i + 1 < slotLeaves.Count)
                {
                    var (s2, n2, c2) = slotLeaves[i + 1];
                    grid.AddRow(col1, new Markup($"[blue][[{Markup.Escape(s2)}]][/] {FormatDeviceName(n2, c2)}"));
                }
                else
                {
                    grid.AddRow(col1, new Text(""));
                }
            }

            busNode.AddNode(grid);
        }
    }

    static void ProcessAddress(XElement addr, TreeNode parent)
    {
        string addrType = addr.Attribute("type")?.Value ?? "";
        string addrValue = addr.Attribute("value")?.Value ?? "";

        if (addrType == "Empty")
            return;

        // Find child device (if any)
        var device = addr.Element("device");
        if (device == null)
            return;

        // Skip back-references
        if (device.Attribute("reference") != null)
            return;

        string? devName = device.Attribute("name")?.Value;
        string? devClass = device.Attribute("classname")?.Value;

        if (devName == null)
            return;

        // Build the display label
        string label;
        if (addrType == "String")
        {
            // IP address: "10.13.30.68 - DeviceName"
            label = $"[cyan]{Markup.Escape(addrValue)}[/] - {FormatDeviceName(devName, devClass)}";
        }
        else if (addrType == "Short")
        {
            // Slot number: "[1] DeviceName"
            label = $"[blue][[{Markup.Escape(addrValue)}]][/] {FormatDeviceName(devName, devClass)}";
        }
        else
        {
            label = $"{Markup.Escape(addrValue)} - {FormatDeviceName(devName, devClass)}";
        }

        var devNode = parent.AddNode(label);

        // Recurse into device's ports → buses → addresses
        foreach (var port in device.Elements("port"))
        {
            foreach (var childBus in port.Elements("bus"))
            {
                ProcessBus(childBus, devNode);
            }
        }
    }

    static string FormatDeviceName(string name, string? classname)
    {
        if (classname == "Unrecognized Device")
            return $"[dim italic]{Markup.Escape(name)}[/]";

        return $"[white]{Markup.Escape(name)}[/]";
    }
}
