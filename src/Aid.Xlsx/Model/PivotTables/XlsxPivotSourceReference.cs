using System.Diagnostics.CodeAnalysis;
using System.Diagnostics;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.PivotTables
{
    /// <summary>
    /// A reference to the source data of <see cref="XlsxPivotCache"/>. The source might exist
    /// or not, that is evaluated during pivot cache record refresh.
    /// </summary>
    internal sealed class XlsxPivotSourceReference : IXlsxPivotSource
    {
        internal XlsxPivotSourceReference(SheetArea area)
        {
            Area = area;
            Name = null;
        }

        internal XlsxPivotSourceReference(string namedRangeOrTable)
        {
            Area = null;
            Name = namedRangeOrTable;
        }

        /// <summary>
        /// Are source data in external workbook defined by a <see cref="Name"/> or by <see cref="Area">cell area</see>.
        /// </summary>
        [MemberNotNullWhen(true, nameof(Name))]
        [MemberNotNullWhen(false, nameof(Area))]
        internal bool UsesName => Name is not null;

        /// <summary>
        /// Book area with the source data. Either this or <see cref="Name"/> is set.
        /// </summary>
        internal SheetArea? Area { get; }

        /// <summary>
        /// Name of a table or a book-scoped named range that contain the source data.
        /// Either this or <see cref="Area"/> is set.
        /// </summary>
        internal string? Name { get; }

        public bool Equals(IXlsxPivotSource? otherSource)
        {
            var other = otherSource as XlsxPivotSourceReference;
            if (other is null)
                return false;

            if (ReferenceEquals(this, other))
                return true;

            return Nullable.Equals(Area, other.Area) && XlsxHelper.NameComparer.Equals(Name, other.Name);
        }

        public override bool Equals(object? obj)
        {
            return obj is IXlsxPivotSource other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return (Area.GetHashCode() * 397) ^ (Name is not null ? XlsxHelper.NameComparer.GetHashCode(Name) : 0);
            }
        }

        /// <summary>
        /// Try to determine actual area of the source reference in the
        /// workbook. Source reference might not be valid in the workbook.
        /// </summary>
        public bool TryGetSource(XlsxWorkbook workbook, out XlsxWorksheet? sheet, out Area? sheetArea)
        {
            if (Name is not null)
            {
                // TODO: Named ranges are currently unusable, so only check tables.
                if (workbook.TryGetTable(Name, out var table))
                {
                    sheet = table.Worksheet;
                    sheetArea = table.Area;
                    return true;
                }

                sheet = null;
                sheetArea = null;
                return false;
            }

            Debug.Assert(Area is not null);
            if (workbook.WorksheetsInternal.TryGetWorksheet(Area.Value.Name, out sheet))
            {
                sheetArea = Area.Value.Area;
                return true;
            }

            sheetArea = default;
            return false;
        }
    }
}
