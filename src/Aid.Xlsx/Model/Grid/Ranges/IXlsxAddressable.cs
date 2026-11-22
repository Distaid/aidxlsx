#nullable disable

namespace Aid.Xlsx.Model.Grid.Ranges;

/// <summary>
/// A very lightweight interface for entities that have an address as
/// a rectangular range.
/// </summary>
public interface IXlsxAddressable
{
    /// <summary>
    ///   Gets an object with the boundaries of this range.
    /// </summary>

    IXlsxRangeAddress RangeAddress { get; }
}
