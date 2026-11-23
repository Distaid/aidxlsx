#nullable disable

namespace Aid.Xlsx.Rendering;

/// <summary>
/// A DPI resolution.
/// </summary>
public readonly struct Dpi(double dpiX, double dpiY)
{
    /// <summary>
    /// Horizontal DPI resolution.
    /// </summary>
    public double X { get; } = dpiX;

    /// <summary>
    /// Vertical DPI resolution.
    /// </summary>
    public double Y { get; } = dpiY;
}
