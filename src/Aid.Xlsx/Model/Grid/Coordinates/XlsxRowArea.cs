using System.Text;
using Aid.Xlsx.Formulas.References;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Model.Grid.Coordinates;

/// <summary>
/// An immutable row address within a workbook.
/// </summary>
internal readonly record struct XlsxRowArea
{
    public XlsxRowArea(string name, int rowNumber)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentException(null, nameof(name));

        if (rowNumber is < XlsxHelper.MinRowNumber or > XlsxHelper.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(rowNumber));

        Name = name;
        RowNumber = rowNumber;
    }

    /// <summary>
    /// Name of the sheet. Sheet may exist or not (e.g. deleted). Never null.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Row number, ranges from 1 to <see cref="XlsxHelper.MaxRowNumber"/>.
    /// </summary>
    public int RowNumber { get; }

    /// <summary>
    /// Get the area of the row.
    /// </summary>
    public SheetArea Area => new(Name, new Area(RowNumber, XlsxHelper.MinColumnNumber, RowNumber, XlsxHelper.MaxColumnNumber));

    public bool Equals(XlsxRowArea other)
    {
        return RowNumber == other.RowNumber && XlsxHelper.SheetComparer.Equals(Name, other.Name);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (XlsxHelper.SheetComparer.GetHashCode(Name) * 397) ^ RowNumber.GetHashCode();
        }
    }

    public override string ToString()
    {
        var name = NameUtils.ShouldQuote(Name.AsSpan()) ? Name.AlwaysEscapeSheetName() : Name;
        return new StringBuilder(name.Length + 1 + 7 + 1 + 7)
            .Append(name).Append('!').Append(RowNumber).Append(':').Append(RowNumber)
            .ToString();
    }
}
