using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.DefinedNames;

namespace Aid.Xlsx.Model.Grid.Coordinates;

/// <summary>
/// A name in a worksheet. Unlike <see cref="IXlsxDefinedName"/>, this is basically only a reference.
/// The actual
/// </summary>
internal readonly struct XlsxName : IEquatable<XlsxName>
{
    /// <summary>
    /// Name of a sheet. If null, the scope is a workbook. The sheet might not exist, e.g. it
    /// is only in a formula. The name of a sheet is not escaped.
    /// </summary>
    public string? SheetName { get; }

    /// <summary>
    /// The defined name in the scope. Case insensitive during comparisons.
    /// </summary>
    public string Name { get; }

    public XlsxName(string sheetName, string name)
    {
        if (string.IsNullOrEmpty(sheetName))
            throw new ArgumentException(null, nameof(sheetName));

        if (name.Any(char.IsWhiteSpace))
            throw new ArgumentException("Name can't contain whitespace.");

        SheetName = sheetName;
        Name = name;
    }

    public XlsxName(string name)
    {
        if (name.Any(char.IsWhiteSpace))
            throw new ArgumentException("Name can't contain whitespace.");

        SheetName = null;
        Name = name;
    }

    public bool Equals(XlsxName other)
    {
        var differentScope = SheetName is null ^ other.SheetName is null;
        if (differentScope)
            return false;

        var bothWorkbookScope = SheetName is null && other.SheetName is null;
        if (bothWorkbookScope)
            return XlsxHelper.NameComparer.Equals(Name, other.Name);

        return XlsxHelper.NameComparer.Equals(Name, other.Name) &&
               XlsxHelper.SheetComparer.Equals(SheetName, other.SheetName);
    }

    public override bool Equals(object? obj)
    {
        return obj is XlsxName other && Equals(other);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = (SheetName is not null ? XlsxHelper.SheetComparer.GetHashCode(SheetName) : 0) * 397;
            hashCode ^= XlsxHelper.NameComparer.GetHashCode(Name);
            return hashCode;
        }
    }

    public override string ToString()
    {
        var isWorkbookScoped = SheetName is null;
        return isWorkbookScoped ? Name : $"{SheetName}!{Name}";
    }
}
