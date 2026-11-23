using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Calculation;

/// <summary>
/// Reference is a collection of cells in the workbook. It's used in formula evaluation.
/// Every reference has at least one cell.
/// </summary>
internal class Reference
{
    public Reference(XlsxRangeAddress area)
    {
        if (!area.IsNormalized)
            throw new ArgumentException("The area must be normalized.", nameof(area));

        Areas = [area];
    }

    /// <summary>
    /// Ctor that reuses parameter to keep allocations low - don't modify the collection after it is passed to ctor.
    /// </summary>
    public Reference(List<XlsxRangeAddress> areas)
    {
        if (areas.Count < 1)
            throw new ArgumentException("At least one area is required.", nameof(areas));

        Areas = areas;
    }

    public Reference(IXlsxRanges ranges)
    {
        if (ranges.Count < 1)
            throw new ArgumentException("At least one range is required.", nameof(ranges));

        var areas = new List<XlsxRangeAddress>(ranges.Count);
        foreach (IXlsxRange range in ranges)
            areas.Add((XlsxRangeAddress)range.RangeAddress);

        Areas = areas;
    }

    /// <summary>
    /// List of areas of the range (at least one). All areas are valid and normalized. Some areas have worksheet and some don't.
    /// </summary>
    internal IReadOnlyList<XlsxRangeAddress> Areas { get; }

    /// <summary>
    /// Get total number of cells coverted by all areas (double counts overlapping areas).
    /// </summary>
    internal int NumberOfCells
    {
        get
        {
            var size = 0;
            for (var i = 0; i < Areas.Count; ++i)
                size += Areas[i].NumberOfCells;

            return size;
        }
    }

    /// <summary>
    /// An iterator over all nonblank cells of the range. Some cells can be iterated
    /// over multiple times (e.g. a union of two ranges with overlapping cells).
    /// </summary>
    public IEnumerable<ScalarValue> GetCellsValues(CalcContext ctx)
    {
        foreach (XlsxRangeAddress area in Areas)
        {
            for (var row = area.FirstAddress.RowNumber; row <= area.LastAddress.RowNumber; ++row)
            {
                for (var column = area.FirstAddress.ColumnNumber; column <= area.LastAddress.ColumnNumber; ++column)
                {
                    ScalarValue cellValue = ctx.GetCellValue(area.Worksheet, row, column);
                    if (!cellValue.IsBlank)
                    {
                        yield return cellValue;
                    }
                }
            }
        }
    }

    public static OneOf<Reference, XlsxError> RangeOp(Reference lhs, Reference rhs, XlsxWorksheet contextWorksheet)
    {
        List<XlsxWorksheet> lhsWorksheets = lhs.Areas.Count == 1
            ? [.. lhs.Areas.Select(a => a.Worksheet).Where(ws => ws is not null).Select(ws => ws!)]
            : [.. lhs.Areas.Select(a => a.Worksheet ?? contextWorksheet).Distinct()];
        if (lhsWorksheets.Count > 1)
            return XlsxError.IncompatibleValue;

        XlsxWorksheet? lhsWorksheet = lhsWorksheets.SingleOrDefault();

        List<XlsxWorksheet> rhsWorksheets = rhs.Areas.Count == 1
            ? [.. rhs.Areas.Select(a => a.Worksheet).Where(ws => ws is not null).Select(ws => ws!)]
            : [.. rhs.Areas.Select(a => a.Worksheet ?? contextWorksheet).Distinct()];
        if (rhsWorksheets.Count > 1)
            return XlsxError.IncompatibleValue;

        XlsxWorksheet? rhsWorksheet = rhsWorksheets.SingleOrDefault();

        if (rhsWorksheet is not null)
        {
            if ((lhsWorksheet ?? contextWorksheet) != rhsWorksheet)
                return XlsxError.IncompatibleValue;
        }

        var minCol = XlsxHelper.MaxColumnNumber;
        var maxCol = 1;
        var minRow = XlsxHelper.MaxRowNumber;
        var maxRow = 1;
        foreach (XlsxRangeAddress area in lhs.Areas.Concat(rhs.Areas))
        {
            // Areas are normalized, so I don't have to check opposite corners
            minRow = Math.Min(minRow, area.FirstAddress.RowNumber);
            maxRow = Math.Max(maxRow, area.LastAddress.RowNumber);
            minCol = Math.Min(minCol, area.FirstAddress.ColumnNumber);
            maxCol = Math.Max(maxCol, area.LastAddress.ColumnNumber);
        }

        XlsxWorksheet? sheet = lhsWorksheet ?? rhsWorksheet;
        return new Reference(new XlsxRangeAddress(
            new XlsxAddress(sheet, minRow, minCol, false, false),
            new XlsxAddress(sheet, maxRow, maxCol, false, false)));
    }

    public static Reference UnionOp(Reference lhs, Reference rhs)
    {
        return new Reference([.. lhs.Areas, .. rhs.Areas]);
    }

    public static OneOf<Reference, XlsxError> Intersect(Reference lhs, Reference rhs, CalcContext ctx)
    {
        var sheets = lhs.Areas.Select(a => a.Worksheet ?? ctx.Worksheet)
            .Concat(rhs.Areas.Select(a => a.Worksheet ?? ctx.Worksheet))
            .Distinct().ToList();
        if (sheets.Count != 1)
            return XlsxError.IncompatibleValue;

        XlsxWorksheet sheet = sheets.Single();
        var intersections = new List<XlsxRangeAddress>();
        foreach (XlsxRangeAddress leftArea in lhs.Areas)
        {
            XlsxRangeAddress intersectedArea = leftArea.WithWorksheet(sheet);
            foreach (XlsxRangeAddress rightArea in rhs.Areas)
            {
                intersectedArea = intersectedArea.Intersection(rightArea.WithWorksheet(sheet));
                if (!intersectedArea.IsValid)
                    break;
            }

            if (intersectedArea.IsValid)
                intersections.Add(intersectedArea);
        }

        return intersections.Count != 0 ? new Reference(intersections) : XlsxError.NullValue;
    }

    /// <summary>
    /// Do an implicit intersection of an address.
    /// </summary>
    /// <param name="formulaAddress"></param>
    /// <returns>An address of the intersection or error if intersection failed.</returns>
    public OneOf<Reference, XlsxError> ImplicitIntersection(IXlsxAddress formulaAddress)
    {
        if (Areas.Count != 1)
            return XlsxError.IncompatibleValue;

        XlsxRangeAddress area = Areas.Single();
        if (area.RowSpan == 1 && area.ColumnSpan == 1)
            return this;

        var column = formulaAddress.ColumnNumber;
        var row = formulaAddress.RowNumber;

        if (area.ColumnSpan == 1 && area.FirstAddress.RowNumber <= row && row <= area.LastAddress.RowNumber)
        {
            var intersection = new XlsxAddress(area.Worksheet, row, area.FirstAddress.ColumnNumber, false, false);
            return new Reference(new XlsxRangeAddress(intersection, intersection));
        }

        if (area.RowSpan == 1 && area.FirstAddress.ColumnNumber <= column && column <= area.LastAddress.ColumnNumber)
        {
            var intersection = new XlsxAddress(area.Worksheet, area.FirstAddress.RowNumber, column, false, false);
            return new Reference(new XlsxRangeAddress(intersection, intersection));
        }

        return XlsxError.IncompatibleValue;
    }

    internal bool IsSingleCell()
    {
        return Areas.Count == 1 && Areas[0].IsSingleCell();
    }

    internal bool TryGetSingleCellValue(out ScalarValue value, CalcContext ctx)
    {
        if (!IsSingleCell())
        {
            value = default;
            return false;
        }

        XlsxRangeAddress area = Areas.Single();
        value = ctx.GetCellValue(area.Worksheet, area.FirstAddress.RowNumber, area.FirstAddress.ColumnNumber);
        return true;
    }

    internal OneOf<Array, XlsxError> ToArray(CalcContext context)
    {
        if (Areas.Count != 1)
            return XlsxError.IncompatibleValue;

        XlsxRangeAddress area = Areas.Single();

        return new ReferenceArray(area, context);
    }

    public OneOf<Array, XlsxError> Apply(Func<ScalarValue, ScalarValue> op, CalcContext context)
    {
        if (Areas.Count != 1)
            return XlsxError.IncompatibleValue;

        XlsxRangeAddress area = Areas.Single();
        var width = area.ColumnSpan;
        var height = area.RowSpan;
        var startColumn = area.FirstAddress.ColumnNumber;
        var startRow = area.FirstAddress.RowNumber;
        ScalarValue[,] data = context.AllocateArray<ScalarValue>(height, width);
        for (int y = 0; y < height; ++y)
        {
            for (int x = 0; x < width; ++x)
            {
                var row = startRow + y;
                var column = startColumn + x;
                ScalarValue cellValue = context.GetCellValue(area.Worksheet, row, column);
                data[y, x] = op(cellValue);
            }
        }

        return new ConstArray(data);
    }
}
