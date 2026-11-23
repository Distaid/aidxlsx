using System.Collections;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Cells;

internal class XlsxCells(XlsxWorkbook workbook, bool usedCellsOnly, XlsxCellsUsedOptions options, Func<XlsxCell, bool>? predicate = null) : IXlsxCells, IEnumerable<XlsxCell>
{
    private readonly XlsxWorkbook _workbook = workbook;
    private readonly List<XlsxRangeAddress> _rangeAddresses = [];
    private readonly bool _usedCellsOnly = usedCellsOnly;
    private readonly Func<XlsxCell, bool> _predicate = predicate ?? (_ => true);
    private readonly XlsxCellsUsedOptions _options = options;

    public XlsxCells(XlsxWorksheet worksheet, bool usedCellsOnly, XlsxCellsUsedOptions options, Func<XlsxCell, bool>? predicate = null)
        : this(worksheet.Workbook, usedCellsOnly, options, predicate)
    {
    }

    #region IEnumerable<XlsxCell> Members

    private IEnumerable<XlsxCell> GetAllCells()
    {
        IEnumerable<IGrouping<XlsxWorksheet?, XlsxRangeAddress>> groupedAddresses = _rangeAddresses.GroupBy(addr => addr.Worksheet);
        foreach (IGrouping<XlsxWorksheet?, XlsxRangeAddress> worksheetGroup in groupedAddresses)
        {
            XlsxWorksheet ws = worksheetGroup.Key!;
            IEnumerable<Point> sheetPoints = worksheetGroup.SelectMany(addr => GetAllCellsInRange(addr))
                .Distinct();
            foreach (Point sheetPoint in sheetPoints)
            {
                XlsxCell c = ws.Cell(sheetPoint.Row, sheetPoint.Column);
                if (_predicate(c))
                    yield return c;
            }
        }
    }

    private static IEnumerable<Point> GetAllCellsInRange(IXlsxRangeAddress rangeAddress)
    {
        if (!rangeAddress.IsValid)
            yield break;

        XlsxRangeAddress normalizedAddress = ((XlsxRangeAddress)rangeAddress).Normalize();
        var minRow = normalizedAddress.FirstAddress.RowNumber;
        var maxRow = normalizedAddress.LastAddress.RowNumber;
        var minColumn = normalizedAddress.FirstAddress.ColumnNumber;
        var maxColumn = normalizedAddress.LastAddress.ColumnNumber;

        for (var ro = minRow; ro <= maxRow; ro++)
        {
            for (var co = minColumn; co <= maxColumn; co++)
            {
                yield return new Point(ro, co);
            }
        }
    }

    private IEnumerable<XlsxCell> GetUsedCells()
    {
        var visitedCells = new HashSet<XlsxAddress>();
        IEnumerable<IGrouping<XlsxWorksheet?, XlsxRangeAddress>> groupedAddresses = _rangeAddresses.GroupBy(addr => addr.Worksheet);
        foreach (IGrouping<XlsxWorksheet?, XlsxRangeAddress> worksheetGroup in groupedAddresses)
        {
            XlsxWorksheet ws = worksheetGroup.Key!;

            IEnumerable<Point> usedCellsCandidates = GetUsedCellsCandidates(ws);

            IOrderedEnumerable<XlsxCell> cells = worksheetGroup.SelectMany(addr => GetUsedCellsInRange(addr, ws, usedCellsCandidates))
                .OrderBy(cell => cell.Address.RowNumber)
                .ThenBy(cell => cell.Address.ColumnNumber);

            visitedCells.Clear();
            foreach (XlsxCell? cell in cells)
            {
                if (visitedCells.Add(cell.Address))
                {
                    yield return cell;
                }
            }
        }
    }

    private IEnumerable<XlsxCell> GetUsedCellsInRange(XlsxRangeAddress rangeAddress, XlsxWorksheet worksheet, IEnumerable<Point> usedCellsCandidates)
    {
        if (!rangeAddress.IsValid)
            yield break;
        XlsxRangeAddress normalizedAddress = rangeAddress.Normalize();
        var minRow = normalizedAddress.FirstAddress.RowNumber;
        var maxRow = normalizedAddress.LastAddress.RowNumber;
        var minColumn = normalizedAddress.FirstAddress.ColumnNumber;
        var maxColumn = normalizedAddress.LastAddress.ColumnNumber;

        IEnumerable<XlsxCell> cellRange = worksheet.Internals.CellsCollection
            .GetCells(minRow, minColumn, maxRow, maxColumn, _predicate);

        foreach (XlsxCell cell in cellRange)
        {
            if (!cell.IsEmpty(_options) && _predicate(cell))
                yield return cell;
        }

        foreach (Point sheetPoint in usedCellsCandidates)
        {
            if (sheetPoint.Row.Between(minRow, maxRow) &&
                sheetPoint.Column.Between(minColumn, maxColumn))
            {
                XlsxCell cell = worksheet.Cell(sheetPoint.Row, sheetPoint.Column);

                if (_predicate(cell))
                    yield return cell;
            }
        }
    }

    private IEnumerable<Point> GetUsedCellsCandidates(XlsxWorksheet worksheet)
    {
        IEnumerable<Point> candidates = [];

        if (_options == XlsxCellsUsedOptions.AllContents)
        {
            return candidates;
        }

        if (_options.HasFlag(XlsxCellsUsedOptions.MergedRanges))
            candidates = candidates.Union(
                worksheet.Internals.MergedRanges.SelectMany<XlsxRange, Point>(r => GetAllCellsInRange(r.RangeAddress)));

        if (_options.HasFlag(XlsxCellsUsedOptions.ConditionalFormats))
            candidates = candidates.Union(
                worksheet.ConditionalFormats.SelectMany<XlsxConditionalFormat, Point>(cf => cf.Ranges.SelectMany(r => GetAllCellsInRange(r.RangeAddress))));

        if (_options.HasFlag(XlsxCellsUsedOptions.DataValidation))
            candidates = candidates.Union(
                worksheet.DataValidations.SelectMany<XlsxDataValidation, Point>(dv => dv.Ranges.SelectMany(r => GetAllCellsInRange(r.RangeAddress))));

        if (_options.HasFlag(XlsxCellsUsedOptions.Sparklines))
            candidates = candidates.Union(
                worksheet.SparklineGroups.SelectMany(sg => sg).Select(sl => Point.FromAddress(sl.Location.Address)));

        return candidates.Distinct();
    }

    public IEnumerator<XlsxCell> GetEnumerator()
    {
        return GetCells().GetEnumerator();
    }

    private IEnumerable<XlsxCell> GetCells()
    {
        return _usedCellsOnly ? GetUsedCells() : GetAllCells();
    }

    #endregion IEnumerable<XlsxCell> Members

    #region IXlsxCells Members

    IEnumerator<IXlsxCell> IEnumerable<IXlsxCell>.GetEnumerator()
    {
        return GetCells().GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public XlsxCellValue Value
    {
        set { this.ForEach<XlsxCell>(c => c.Value = value); }
    }

    public IXlsxCells Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        this.ForEach<XlsxCell>(c => c.Clear(clearOptions));
        return this;
    }

    public void DeleteComments()
    {
        this.ForEach<XlsxCell>(c => c.DeleteComment());
    }

    public void DeleteSparklines()
    {
        this.ForEach<XlsxCell>(c => c.DeleteSparkline());
    }

    public string FormulaA1
    {
        set { this.ForEach<XlsxCell>(c => c.FormulaA1 = value); }
    }

    public string FormulaR1C1
    {
        set { this.ForEach<XlsxCell>(c => c.FormulaR1C1 = value); }
    }

    public IXlsxStyle Style
    {
        get => Format;
        set => Format.SetStyle(value);
    }

    internal XlsxCellFormat Format
    {
        get
        {
            // For backwards compatibility, the sheet is considered the inner style. A terrible
            // choice, but it is what it is.
            XlsxWorksheet? sheet = _rangeAddresses.Select(x => x.Worksheet).FirstOrDefault(x => x is not null);
            SheetArea[] areas = [.. _rangeAddresses.Select(SheetArea.From)];
            return XlsxCellFormat.ForCells(_workbook, areas, sheet);
        }
    }

    #endregion IXlsxCells Members

    public void Add(XlsxRangeAddress rangeAddress)
    {
        _rangeAddresses.Add(rangeAddress);
    }

    public void Add(XlsxCell cell)
    {
        Add(new XlsxRangeAddress(cell.Address, cell.Address));
    }

    public void Select()
    {
        foreach (XlsxCell cell in this)
            cell.Select();
    }
}
