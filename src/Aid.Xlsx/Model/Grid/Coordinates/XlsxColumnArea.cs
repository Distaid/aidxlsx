using System.Text;
using Aid.Xlsx.Formulas.References;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Model.Grid.Coordinates;

/// <summary>
/// An immutable column address within a workbook.
/// </summary>
internal readonly record struct XlsxColumnArea
{
    public XlsxColumnArea(string name, int columnNumber)
    {
        ArgumentException.ThrowIfNullOrEmpty(name);

        if (columnNumber is < XlsxHelper.MinColumnNumber or > XlsxHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnNumber));

        Name = name;
        ColumNumber = columnNumber;
    }

    /// <summary>
    /// Name of the sheet. Sheet may exist or not (e.g. deleted). Never null.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Column number, ranges from 1 to <see cref="XlsxHelper.MaxColumnNumber"/>.
    /// </summary>
    public int ColumNumber { get; }

    public SheetArea Area => new(Name, new Area(XlsxHelper.MinRowNumber, ColumNumber, XlsxHelper.MaxRowNumber, ColumNumber));

    public bool Equals(XlsxColumnArea other)
    {
        return ColumNumber == other.ColumNumber && XlsxHelper.SheetComparer.Equals(Name, other.Name);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            return (XlsxHelper.SheetComparer.GetHashCode(Name) * 397) ^ ColumNumber.GetHashCode();
        }
    }

    public override string ToString()
    {
        var name = NameUtils.ShouldQuote(Name.AsSpan()) ? Name.AlwaysEscapeSheetName() : Name;
        var column = XlsxHelper.GetColumnLetterFromNumber(ColumNumber);
        return new StringBuilder(name.Length + 2 + 2 * column.Length)
            .Append(name).Append('!').Append(column).Append(':').Append(column)
            .ToString();
    }
}
