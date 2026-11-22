using Aid.Xlsx.Extensions;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.Coordinates;

/// <summary>
/// <see cref="XlsxReference"/> that includes a sheet. It can represent cell
/// (<c>'Sheet one'!A$1</c>), area (<c>Sheet1!A4:$G$5</c>), row span (<c>'Sales Q1'!4:10</c>)
/// and col span (<c>Sales!G:H</c>).
/// </summary>
/// <param name="Sheet">
/// Name of a sheet. Unescaped, so it doesn't include quotes. Note that sheet might not exist.
/// </param>
/// <param name="Reference">
/// Referenced area in the sheet. Can be in A1 or R1C1.
/// </param>
internal readonly record struct XlsxSheetReference(string Sheet, XlsxReference Reference)
{
    public bool Equals(XlsxSheetReference other)
    {
        return XlsxHelper.SheetComparer.Equals(Sheet, other.Sheet) &&
               Reference.Equals(other.Reference);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = 2122234362;
            hashCode = hashCode * -1521134295 + XlsxHelper.SheetComparer.GetHashCode(Sheet);
            hashCode = hashCode * -1521134295 + Reference.GetHashCode();
            return hashCode;
        }
    }

    internal string GetA1()
    {
        return Sheet.EscapeSheetName() + '!' + Reference.GetA1();
    }
}
