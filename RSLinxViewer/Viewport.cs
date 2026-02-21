using Spectre.Console.Rendering;

namespace RSLinxViewer;

/// <summary>
/// Wraps an IRenderable and clips output to a vertical viewport.
/// Only renders lines from offset to offset+height.
/// After Render(), TotalLines contains the full content height.
/// </summary>
sealed class Viewport : IRenderable
{
    readonly IRenderable _inner;
    readonly int _offset;
    readonly int _height;

    /// <summary>Total lines in the inner renderable (set after Render).</summary>
    public int TotalLines { get; private set; }

    public Viewport(IRenderable inner, int offset, int height)
    {
        _inner = inner;
        _offset = Math.Max(0, offset);
        _height = Math.Max(1, height);
    }

    public Measurement Measure(RenderOptions options, int maxWidth)
    {
        return _inner.Measure(options, maxWidth);
    }

    public IEnumerable<Segment> Render(RenderOptions options, int maxWidth)
    {
        // Eagerly collect all segments to count total lines
        var allSegments = _inner.Render(options, maxWidth).ToList();

        int totalLines = 1;
        foreach (var seg in allSegments)
        {
            if (seg.IsLineBreak)
                totalLines++;
        }
        TotalLines = totalLines;

        // Clamp offset so we can't scroll past the end
        int maxOffset = Math.Max(0, totalLines - _height);
        int clampedOffset = Math.Min(_offset, maxOffset);
        int endLine = clampedOffset + _height;

        // Emit only the visible window
        int currentLine = 0;
        foreach (var segment in allSegments)
        {
            if (segment.IsLineBreak)
            {
                if (currentLine >= clampedOffset && currentLine < endLine)
                    yield return segment;
                currentLine++;
                if (currentLine >= endLine)
                    yield break;
                continue;
            }

            if (currentLine >= clampedOffset && currentLine < endLine)
                yield return segment;
        }
    }
}
