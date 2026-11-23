using System.Collections;
using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Grid.Columns;

/// <summary>
/// Create a new instance of <see cref="XlsxColumns"/>.
/// </summary>
/// <param name="workbook">Workbook to which all columns belong.</param>
/// <param name="worksheet">If worksheet is specified it means that the created instance represents
/// all columns on a worksheet so changing its width will affect all columns.</param>
/// <param name="defaultStyleSheet">A sheet with a default style to use when initializing child entries.</param>
/// <param name="lazyEnumerable">A predefined enumerator of <see cref="XlsxColumn"/> to support lazy initialization.</param>
internal class XlsxColumns(XlsxWorkbook workbook, XlsxWorksheet? worksheet, XlsxWorksheet? defaultStyleSheet = null, IEnumerable<XlsxColumn>? lazyEnumerable = null) : IXlsxColumns
{
    private readonly List<XlsxColumn> _columnsCollection = [];

    private readonly XlsxWorkbook _workbook = workbook;
    private readonly XlsxWorksheet? _worksheet = worksheet;
    private readonly XlsxWorksheet? _defaultStyleSheet = defaultStyleSheet;

    /// <summary>
    /// This object represents all columns of the worksheet, even non-materialized ones.
    /// </summary>
    [MemberNotNullWhen(true, nameof(_worksheet))]
    private bool AllColumnsOfSheet => _worksheet is not null;

    private bool IsMaterialized => _lazyEnumerable == null;

    private IEnumerable<XlsxColumn>? _lazyEnumerable = lazyEnumerable;

    private IEnumerable<XlsxColumn> Columns => _lazyEnumerable ?? _columnsCollection.AsEnumerable();

    #region IXlsxColumns Members

    public IEnumerator<IXlsxColumn> GetEnumerator()
    {
        return Columns.Cast<IXlsxColumn>().OrderBy(r => r.ColumnNumber()).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public double Width
    {
        set
        {
            Columns.ForEach(c => c.Width = value);

            if (!AllColumnsOfSheet) return;

            _worksheet.ColumnWidth = value;
            _worksheet.Internals.ColumnsCollection.ForEach(c => c.Value.Width = value);
        }
    }

    public void Delete()
    {
        if (AllColumnsOfSheet)
        {
            _worksheet.Internals.ColumnsCollection.Clear();
            _worksheet.Internals.CellsCollection.Clear();
        }
        else
        {
            var toDelete = new Dictionary<IXlsxWorksheet, List<int>>();
            foreach (XlsxColumn c in Columns)
            {
                if (!toDelete.TryGetValue(c.Worksheet, out List<int>? list))
                {
                    list = [];
                    toDelete.Add(c.Worksheet, list);
                }

                list.Add(c.ColumnNumber());
            }

            foreach (KeyValuePair<IXlsxWorksheet, List<int>> kp in toDelete)
            {
                foreach (int c in kp.Value.OrderByDescending(c => c))
                    kp.Key.Column(c).Delete();
            }
        }
    }

    public IXlsxColumns AdjustToContents()
    {
        Columns.ForEach(c => c.AdjustToContents());
        return this;
    }

    public IXlsxColumns AdjustToContents(int startRow)
    {
        Columns.ForEach(c => c.AdjustToContents(startRow));
        return this;
    }

    public IXlsxColumns AdjustToContents(int startRow, int endRow)
    {
        Columns.ForEach(c => c.AdjustToContents(startRow, endRow));
        return this;
    }

    public IXlsxColumns AdjustToContents(double minWidth, double maxWidth)
    {
        Columns.ForEach(c => c.AdjustToContents(minWidth, maxWidth));
        return this;
    }

    public IXlsxColumns AdjustToContents(int startRow, double minWidth, double maxWidth)
    {
        Columns.ForEach(c => c.AdjustToContents(startRow, minWidth, maxWidth));
        return this;
    }

    public IXlsxColumns AdjustToContents(int startRow, int endRow, double minWidth, double maxWidth)
    {
        Columns.ForEach(c => c.AdjustToContents(startRow, endRow, minWidth, maxWidth));
        return this;
    }

    public void Hide()
    {
        Columns.ForEach(c => c.Hide());
    }

    public void Unhide()
    {
        Columns.ForEach(c => c.Unhide());
    }

    public void Group()
    {
        Group(false);
    }

    public void Group(int outlineLevel)
    {
        Group(outlineLevel, false);
    }

    public void Ungroup()
    {
        Ungroup(false);
    }

    public void Group(bool collapse)
    {
        Columns.ForEach(c => c.Group(collapse));
    }

    public void Group(int outlineLevel, bool collapse)
    {
        Columns.ForEach(c => c.Group(outlineLevel, collapse));
    }

    public void Ungroup(bool ungroupFromAll)
    {
        Columns.ForEach(c => c.Ungroup(ungroupFromAll));
    }

    public void Collapse()
    {
        Columns.ForEach(c => c.Collapse());
    }

    public void Expand()
    {
        Columns.ForEach(c => c.Expand());
    }

    public IXlsxCells Cells()
    {
        var cells = new XlsxCells(_workbook, false, XlsxCellsUsedOptions.All);
        foreach (XlsxColumn container in Columns)
            cells.Add(container.RangeAddress);
        return cells;
    }

    public IXlsxCells CellsUsed()
    {
        var cells = new XlsxCells(_workbook, true, XlsxCellsUsedOptions.All);
        foreach (XlsxColumn container in Columns)
            cells.Add(container.RangeAddress);
        return cells;
    }

    public IXlsxCells CellsUsed(bool includeFormats)
    {
        return CellsUsed(includeFormats
            ? XlsxCellsUsedOptions.All
            : XlsxCellsUsedOptions.AllContents);
    }

    public IXlsxCells CellsUsed(XlsxCellsUsedOptions options)
    {
        var cells = new XlsxCells(_workbook, true, options);
        foreach (XlsxColumn container in Columns)
            cells.Add(container.RangeAddress);
        return cells;
    }

    /// <summary>
    ///   Adds a vertical page break after this column.
    /// </summary>
    public IXlsxColumns AddVerticalPageBreaks()
    {
        foreach (XlsxColumn col in Columns)
            col.Worksheet.PageSetup.AddVerticalPageBreak(col.ColumnNumber());
        return this;
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
            if (AllColumnsOfSheet)
            {
                return XlsxCellFormat.ForWorksheet(_worksheet);
            }

            return XlsxCellFormat.ForColumns(_workbook, _defaultStyleSheet, Columns);
        }
    }

    #endregion IXlsxColumns Members

    public void Add(XlsxColumn column)
    {
        Materialize();
        _columnsCollection.Add(column);
    }

    public void CollapseOnly()
    {
        Columns.ForEach(c => c.Collapsed = true);
    }

    public IXlsxColumns Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
    {
        Columns.ForEach(c => c.Clear(clearOptions));
        return this;
    }

    public void Select()
    {
        foreach (IXlsxColumn range in this)
            range.Select();
    }

    private void Materialize()
    {
        if (IsMaterialized)
            return;

        _columnsCollection.AddRange(Columns);
        _lazyEnumerable = null;
    }
}
