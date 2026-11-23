using static Aid.Xlsx.Model.Protection.XlsxProtectionAlgorithm;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.DataImport.RowReaders;
using Aid.Xlsx.Infrastructure.Caching;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.DefinedNames;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.Drawings.Charts;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Index;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.Hyperlinks;
using Aid.Xlsx.Model.PageSetup;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Sparklines;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.Worksheets;

internal class XlsxWorksheet : XlsxRangeBase, IXlsxWorksheet, IXlsxFormatContainer
{
    #region Fields

    private readonly Dictionary<Int32, Int32> _columnOutlineCount = [];
    private readonly Dictionary<Int32, Int32> _rowOutlineCount = [];
    private readonly XlsxRangeFactory _rangeFactory;
    private readonly XlsxRangeRepository _rangeRepository;
    private readonly List<IXlsxRangeIndex> _rangeIndices;
    private readonly XlsxRanges _selectedRanges;

    internal Int32 ZOrder = 1;
    private String _name;
    internal Int32 _position;

    private Double _rowHeight;
    private Boolean _tabActive;
    private XlsxSheetProtection _protection;

    /// <summary>
    /// Fake address to be used everywhere the invalid address is needed.
    /// </summary>
    internal readonly XlsxAddress InvalidAddress;

    #endregion Fields

    #region Constructor

    public XlsxWorksheet(String sheetName, XlsxWorkbook workbook, UInt32 sheetId)
        : base(
            new XlsxRangeAddress(
                new XlsxAddress(null, XlsxHelper.MinRowNumber, XlsxHelper.MinColumnNumber, false, false),
                new XlsxAddress(null, XlsxHelper.MaxRowNumber, XlsxHelper.MaxColumnNumber, false, false))
            )
    {
        Workbook = workbook;
        SheetId = sheetId;
        InvalidAddress = new XlsxAddress(this, 0, 0, false, false);

        var firstAddress = new XlsxAddress(this, RangeAddress.FirstAddress.RowNumber, RangeAddress.FirstAddress.ColumnNumber,
                                               RangeAddress.FirstAddress.FixedRow, RangeAddress.FirstAddress.FixedColumn);
        var lastAddress = new XlsxAddress(this, RangeAddress.LastAddress.RowNumber, RangeAddress.LastAddress.ColumnNumber,
                                              RangeAddress.LastAddress.FixedRow, RangeAddress.LastAddress.FixedColumn);
        RangeAddress = new XlsxRangeAddress(firstAddress, lastAddress);
        _rangeFactory = new XlsxRangeFactory(this);
        _rangeRepository = new XlsxRangeRepository(workbook, _rangeFactory.Create);
        _rangeIndices = [];

        Pictures = new XlsxPictures(this);
        DefinedNames = new XlsxDefinedNames(this);
        SheetView = new XlsxSheetView(this);
        Tables = [];
        Hyperlinks = new XlsxHyperlinks(this);
        DataValidations = new XlsxDataValidations(this);
        PivotTables = new XlsxPivotTables(this);
        _protection = new XlsxSheetProtection(DefaultProtectionAlgorithm);
        AutoFilter = new XlsxAutoFilter();
        ConditionalFormats = new XlsxConditionalFormats(this);
        SparklineGroupsInternal = new XlsxSparklineGroups(this);
        Internals = new XlsxWorksheetInternals
        {
            CellsCollection = new XlsxCellsCollection(this),
            ColumnsCollection = [],
            RowsCollection = [],
            MergedRanges = new XlsxRanges(this)
        };
        PageSetup = new XlsxPageSetup((XlsxPageSetup)workbook.PageOptions, this);
        Outline = new XlsxOutline(workbook.Outline);
        _columnWidth = workbook.ColumnWidth;
        _rowHeight = workbook.RowHeight;
        RowHeightChanged = Math.Abs(workbook.RowHeight - XlsxWorkbook.DefaultRowHeight) > XlsxHelper.Epsilon;

        XlsxHelper.ValidateSheetName(sheetName);
        _name = sheetName;
        Charts = new XlsxCharts();
        ShowFormulas = workbook.ShowFormulas;
        ShowGridLines = workbook.ShowGridLines;
        ShowOutlineSymbols = workbook.ShowOutlineSymbols;
        ShowRowColHeaders = workbook.ShowRowColHeaders;
        ShowRuler = workbook.ShowRuler;
        ShowWhiteSpace = workbook.ShowWhiteSpace;
        ShowZeros = workbook.ShowZeros;
        RightToLeft = workbook.RightToLeft;
        TabColor = XlsxColor.Automatic;
        _selectedRanges = new XlsxRanges(this);

        Author = workbook.Author;
    }

    #endregion Constructor

    internal SheetArea Area => new(Name, Aid.Xlsx.Model.Grid.Coordinates.Area.Full);

    [Obsolete($"Use {nameof(DefinedNames)} instead.")]
    IXlsxDefinedNames IXlsxWorksheet.NamedRanges => DefinedNames;

    IXlsxDefinedNames IXlsxWorksheet.DefinedNames => DefinedNames;

    internal XlsxDefinedNames DefinedNames { get; }

    public override XlsxRangeType RangeType
    {
        get { return XlsxRangeType.Worksheet; }
    }

    /// <summary>
    /// Reference to a VML that contains notes, forms controls and so on. All such things are generally unified into
    /// a single legacy VML file, set during load/save.
    /// </summary>
    public string? LegacyDrawingId;

    private Double _columnWidth;

    public XlsxWorksheetInternals Internals { get; }

    internal XlsxSparklineGroups SparklineGroupsInternal { get; }

    public XlsxRangeFactory RangeFactory
    {
        get { return _rangeFactory; }
    }

    internal Boolean RowHeightChanged { get; set; }

    internal Boolean ColumnWidthChanged { get; set; }

    /// <summary>
    /// <para>
    /// The id of a sheet that is unique across the workbook, kept across load/save.
    /// The ids of sheets are not reused. That is important for referencing the sheet
    /// range/point through sheetId. If sheetIds were reused, references would refer
    /// to the wrong sheet after the original sheetId was reused. Excel also doesn't
    /// reuse sheetIds.
    /// </para>
    /// <para>
    /// Referencing sheet through non-reused sheetIds means that reference can survive
    /// sheet renaming without any changes. Always &gt; 0 (Excel will crash on 0).
    /// </para>
    /// </summary>
    internal UInt32 SheetId { get; set; }

    /// <summary>
    /// A cached <c>r:id</c> of the sheet from the file. If the sheet is a new one (not
    /// yet saved), the value is null until workbook is saved. Use <see cref="SheetId"/>
    /// instead is possible. Mostly for removing deleted sheet parts during save.
    /// </summary>
    internal String? RelId { get; set; }

    public XlsxDataValidations DataValidations { get; private set; }

    public IXlsxCharts Charts { get; private set; }

    public XlsxSheetProtection Protection
    {
        get => _protection;
        set => _protection = value.Clone().CastTo<XlsxSheetProtection>();
    }

    public XlsxAutoFilter AutoFilter { get; private set; }

    public bool IsDeleted { get; private set; }

    #region IXlsxFormatContainer

    /// <remarks>
    /// Format of a worksheet or <c>null</c> if worksheet has no format. In OOXML, worksheet
    /// doesn't actually has a format by itself, so this is mostly virtual. If all columns have
    /// format during the load, it is set to the format of the last column in a sheet (XFD).
    /// </remarks>
    /// <inheritdoc cref="IXlsxFormatContainer.FormatValue"/>
    public XlsxCellFormatValue? FormatValue { get; set; }

    internal override XlsxCellFormat Format => XlsxCellFormat.ForWorksheet(this);

    internal XlsxCellFormatValue GetFormat()
    {
        return FormatValue ?? Workbook.Styles.DefaultFormat;
    }

    #endregion

    #region IXlsxWorksheet Members

    public XlsxWorkbook Workbook { get; }

    public Double ColumnWidth
    {
        get { return _columnWidth; }
        set
        {
            ColumnWidthChanged = true;
            _columnWidth = value;
        }
    }

    public Double RowHeight
    {
        get { return _rowHeight; }
        set
        {
            RowHeightChanged = true;
            _rowHeight = value;
        }
    }

    public String Name
    {
        get { return _name; }
        set
        {
            if (_name == value) return;

            XlsxHelper.ValidateSheetName(value);

            Workbook.WorksheetsInternal.Rename(_name, value);
            _name = value;
        }
    }

    public Int32 Position
    {
        get { return _position; }
        set
        {
            if (value > Workbook.WorksheetsInternal.Count + Workbook.UnsupportedSheets.Count + 1)
                throw new ArgumentOutOfRangeException(nameof(value), "Index must be equal or less than the number of worksheets + 1.");

            if (value < _position)
            {
                Workbook.WorksheetsInternal
                    .Where<XlsxWorksheet>(w => w.Position >= value && w.Position < _position)
                    .ForEach(w => w._position += 1);
            }

            if (value > _position)
            {
                Workbook.WorksheetsInternal
                    .Where<XlsxWorksheet>(w => w.Position <= value && w.Position > _position)
                    .ForEach(w => (w)._position -= 1);
            }

            _position = value;
        }
    }

    public IXlsxPageSetup PageSetup { get; private set; }

    public IXlsxOutline Outline { get; private set; }

    IXlsxRow? IXlsxWorksheet.FirstRowUsed()
    {
        return FirstRowUsed();
    }

    IXlsxRow? IXlsxWorksheet.FirstRowUsed(XlsxCellsUsedOptions options)
    {
        return FirstRowUsed(options);
    }

    IXlsxRow? IXlsxWorksheet.LastRowUsed()
    {
        return LastRowUsed();
    }

    IXlsxRow? IXlsxWorksheet.LastRowUsed(XlsxCellsUsedOptions options)
    {
        return LastRowUsed(options);
    }

    IXlsxColumn IXlsxWorksheet.LastColumn()
    {
        return LastColumn();
    }

    IXlsxColumn IXlsxWorksheet.FirstColumn()
    {
        return FirstColumn();
    }

    IXlsxRow IXlsxWorksheet.FirstRow()
    {
        return FirstRow();
    }

    IXlsxRow IXlsxWorksheet.LastRow()
    {
        return LastRow();
    }

    IXlsxColumn? IXlsxWorksheet.FirstColumnUsed()
    {
        return FirstColumnUsed();
    }

    IXlsxColumn? IXlsxWorksheet.FirstColumnUsed(XlsxCellsUsedOptions options)
    {
        return FirstColumnUsed(options);
    }

    IXlsxColumn? IXlsxWorksheet.LastColumnUsed()
    {
        return LastColumnUsed();
    }

    IXlsxColumn? IXlsxWorksheet.LastColumnUsed(XlsxCellsUsedOptions options)
    {
        return LastColumnUsed(options);
    }

    public IXlsxColumns Columns()
    {
        var columnMap = new HashSet<Int32>();

        columnMap.UnionWith(Internals.CellsCollection.ColumnsUsedKeys);
        columnMap.UnionWith(Internals.ColumnsCollection.Keys);

        var retVal = new XlsxColumns(Workbook, this, this, columnMap.Select(Column));
        return retVal;
    }

    public IXlsxColumns Columns(String columns)
    {
        var retVal = new XlsxColumns(Workbook, null, this);
        var columnPairs = columns.Split(',');
        foreach (string tPair in columnPairs.Select(pair => pair.Trim()))
        {
            String firstColumn;
            String lastColumn;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                var columnRange = XlsxHelper.SplitRange(tPair);
                firstColumn = columnRange[0];
                lastColumn = columnRange[1];
            }
            else
            {
                firstColumn = tPair;
                lastColumn = tPair;
            }

            if (Int32.TryParse(firstColumn, out int tmp))
            {
                foreach (IXlsxColumn col in Columns(Int32.Parse(firstColumn), Int32.Parse(lastColumn)))
                    retVal.Add((XlsxColumn)col);
            }
            else
            {
                foreach (IXlsxColumn col in Columns(firstColumn, lastColumn))
                    retVal.Add((XlsxColumn)col);
            }
        }
        return retVal;
    }

    public IXlsxColumns Columns(String firstColumn, String lastColumn)
    {
        return Columns(XlsxHelper.GetColumnNumberFromLetter(firstColumn),
                       XlsxHelper.GetColumnNumberFromLetter(lastColumn));
    }

    public IXlsxColumns Columns(Int32 firstColumn, Int32 lastColumn)
    {
        var retVal = new XlsxColumns(Workbook, null, this);

        for (int co = firstColumn; co <= lastColumn; co++)
            retVal.Add(Column(co));
        return retVal;
    }

    public IXlsxRows Rows()
    {
        var rowMap = new HashSet<Int32>();

        rowMap.UnionWith(Internals.CellsCollection.RowsUsedKeys);
        rowMap.UnionWith(Internals.RowsCollection.Keys);

        var retVal = new XlsxRows(Workbook, this, this, rowMap.Select(Row));
        return retVal;
    }

    public IXlsxRows Rows(String rows)
    {
        var retVal = new XlsxRows(Workbook, null, this);
        var rowPairs = rows.Split(',');
        foreach (string tPair in rowPairs.Select(pair => pair.Trim()))
        {
            String firstRow;
            String lastRow;
            if (tPair.Contains(':') || tPair.Contains('-'))
            {
                var rowRange = XlsxHelper.SplitRange(tPair);
                firstRow = rowRange[0];
                lastRow = rowRange[1];
            }
            else
            {
                firstRow = tPair;
                lastRow = tPair;
            }

            Rows(Int32.Parse(firstRow), Int32.Parse(lastRow))
                .ForEach(row => retVal.Add((XlsxRow)row));
        }
        return retVal;
    }

    public IXlsxRows Rows(Int32 firstRow, Int32 lastRow)
    {
        var retVal = new XlsxRows(Workbook, null, this);

        for (int ro = firstRow; ro <= lastRow; ro++)
            retVal.Add(Row(ro));
        return retVal;
    }

    IXlsxRow IXlsxWorksheet.Row(Int32 row)
    {
        return Row(row);
    }

    IXlsxColumn IXlsxWorksheet.Column(Int32 column)
    {
        return Column(column);
    }

    IXlsxColumn IXlsxWorksheet.Column(String column)
    {
        return Column(column);
    }

    IXlsxCell IXlsxWorksheet.Cell(int row, int column)
    {
        return Cell(row, column);
    }

    IXlsxCell IXlsxWorksheet.Cell(string cellAddressInRange)
    {
        return Cell(cellAddressInRange) ?? throw new ArgumentException($"'{cellAddressInRange}' is not A1 address or workbook named range.");
    }

    IXlsxCell IXlsxWorksheet.Cell(int row, string column)
    {
        return Cell(row, column);
    }

    IXlsxCell IXlsxWorksheet.Cell(IXlsxAddress cellAddressInRange)
    {
        return Cell(cellAddressInRange);
    }

    IXlsxRange IXlsxWorksheet.Range(IXlsxRangeAddress rangeAddress)
    {
        return Range(rangeAddress);
    }

    IXlsxRange IXlsxWorksheet.Range(string rangeAddress)
    {
        return Range(rangeAddress) ?? throw new ArgumentException($"'{rangeAddress}' is not A1 address or named range.");
    }

    IXlsxRange IXlsxWorksheet.Range(IXlsxCell firstCell, IXlsxCell lastCell)
    {
        return Range(firstCell, lastCell);
    }

    IXlsxRange IXlsxWorksheet.Range(string firstCellAddress, string lastCellAddress)
    {
        return Range(firstCellAddress, lastCellAddress);
    }

    IXlsxRange IXlsxWorksheet.Range(IXlsxAddress firstCellAddress, IXlsxAddress lastCellAddress)
    {
        return Range(firstCellAddress, lastCellAddress);
    }

    IXlsxRange IXlsxWorksheet.Range(int firstCellRow, int firstCellColumn, int lastCellRow, int lastCellColumn)
    {
        return Range(firstCellRow, firstCellColumn, lastCellRow, lastCellColumn);
    }

    IXlsxRanges IXlsxWorksheet.Ranges(String ranges) => Ranges(ranges);

    public IXlsxWorksheet CollapseRows()
    {
        Enumerable.Range(1, 8).ForEach(i => CollapseRows(i));
        return this;
    }

    public IXlsxWorksheet CollapseColumns()
    {
        Enumerable.Range(1, 8).ForEach(i => CollapseColumns(i));
        return this;
    }

    public IXlsxWorksheet ExpandRows()
    {
        Enumerable.Range(1, 8).ForEach(i => ExpandRows(i));
        return this;
    }

    public IXlsxWorksheet ExpandColumns()
    {
        Enumerable.Range(1, 8).ForEach(i => ExpandColumns(i));
        return this;
    }

    public IXlsxWorksheet CollapseRows(Int32 outlineLevel)
    {
        if (outlineLevel < 1 || outlineLevel > 8)
            throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Outline level must be between 1 and 8.");

        Internals.RowsCollection.Values.Where(r => r.OutlineLevel == outlineLevel).ForEach(r => r.Collapse());
        return this;
    }

    public IXlsxWorksheet CollapseColumns(Int32 outlineLevel)
    {
        if (outlineLevel < 1 || outlineLevel > 8)
            throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Outline level must be between 1 and 8.");

        Internals.ColumnsCollection.Values.Where(c => c.OutlineLevel == outlineLevel).ForEach(c => c.Collapse());
        return this;
    }

    public IXlsxWorksheet ExpandRows(Int32 outlineLevel)
    {
        if (outlineLevel < 1 || outlineLevel > 8)
            throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Outline level must be between 1 and 8.");

        Internals.RowsCollection.Values.Where(r => r.OutlineLevel == outlineLevel).ForEach(r => r.Expand());
        return this;
    }

    public IXlsxWorksheet ExpandColumns(Int32 outlineLevel)
    {
        if (outlineLevel < 1 || outlineLevel > 8)
            throw new ArgumentOutOfRangeException(nameof(outlineLevel), "Outline level must be between 1 and 8.");

        Internals.ColumnsCollection.Values.Where(c => c.OutlineLevel == outlineLevel).ForEach(c => c.Expand());
        return this;
    }

    public void Delete()
    {
        IsDeleted = true;
        Workbook.DefinedNamesInternal.OnWorksheetDeleted(Name);
        Workbook.NotifyWorksheetDeleting(this);
        Workbook.WorksheetsInternal.Delete(Name);
    }


    [Obsolete($"Used {nameof(DefinedName)} instead.")]
    IXlsxDefinedName IXlsxWorksheet.NamedRange(String name) => DefinedName(name);

    IXlsxDefinedName IXlsxWorksheet.DefinedName(String name) => DefinedName(name);

    internal XlsxDefinedName DefinedName(String name)
    {
        return DefinedNames.DefinedName(name);
    }

    IXlsxSheetView IXlsxWorksheet.SheetView { get => SheetView; }

    public XlsxSheetView SheetView { get; private set; }

    IXlsxTables IXlsxWorksheet.Tables => Tables;

    internal XlsxTables Tables { get; }

    public IXlsxTable Table(Int32 index)
    {
        return Tables.Table(index);
    }

    public IXlsxTable Table(String name)
    {
        return Tables.Table(name);
    }

    public IXlsxWorksheet CopyTo(String newSheetName)
    {
        return CopyTo(Workbook, newSheetName, Workbook.WorksheetsInternal.Count + 1);
    }

    public IXlsxWorksheet CopyTo(String newSheetName, Int32 position)
    {
        return CopyTo(Workbook, newSheetName, position);
    }

    public IXlsxWorksheet CopyTo(XlsxWorkbook workbook)
    {
        return CopyTo(workbook, Name, workbook.WorksheetsInternal.Count + 1);
    }

    public IXlsxWorksheet CopyTo(XlsxWorkbook workbook, String newSheetName)
    {
        return CopyTo(workbook, newSheetName, workbook.WorksheetsInternal.Count + 1);
    }

    public IXlsxWorksheet CopyTo(XlsxWorkbook workbook, String newSheetName, Int32 position)
    {
        if (this.IsDeleted)
            throw new InvalidOperationException($"`{this.Name}` has been deleted and cannot be copied.");

        var targetSheet = (XlsxWorksheet)workbook.WorksheetsInternal.Add(newSheetName, position);
        Internals.ColumnsCollection.ForEach(kp => kp.Value.CopyTo(targetSheet.Column(kp.Key)));
        Internals.RowsCollection.ForEach(kp => kp.Value.CopyTo(targetSheet.Row(kp.Key)));
        Internals.CellsCollection.GetCells().ForEach(c => targetSheet.Cell(c.Address).CopyFrom(c, XlsxCellCopyOptions.Values | XlsxCellCopyOptions.Styles));
        foreach (var dataValidation in DataValidations)
        {
            targetSheet.DataValidations.CopyFrom(dataValidation);
        }

        targetSheet.Visibility = Visibility;
        targetSheet.ColumnWidth = ColumnWidth;
        targetSheet.ColumnWidthChanged = ColumnWidthChanged;
        targetSheet.RowHeight = RowHeight;
        targetSheet.RowHeightChanged = RowHeightChanged;
        if (FormatValue is not null)
            targetSheet.FormatValue = workbook.Styles.GetRegisteredCellFormat(FormatValue);

        targetSheet.PageSetup = new XlsxPageSetup((XlsxPageSetup)PageSetup, targetSheet);
        ((XlsxHeaderFooter)targetSheet.PageSetup.Header).Changed = true;
        ((XlsxHeaderFooter)targetSheet.PageSetup.Footer).Changed = true;
        targetSheet.Outline = new XlsxOutline(Outline);
        targetSheet.SheetView = new XlsxSheetView(targetSheet, SheetView);
        targetSheet.SelectedRanges.RemoveAll();

        foreach (var picture in Pictures)
            picture.CopyTo(targetSheet);

        Tables.ForEach<XlsxTable>(t => t.CopyTo(targetSheet, false));
        DefinedNames.ForEach<XlsxDefinedName>(nr => nr.CopyTo(targetSheet)); // Names must modify table references, so keep the order.
        PivotTables.ForEach<XlsxPivotTable>(pt => pt.CopyTo(targetSheet.Cell(pt.TargetCell.Address.CastTo<XlsxAddress>().WithoutWorksheet())));
        foreach (var cf in ConditionalFormats)
            cf.CopyTo(targetSheet);

        // Sparklines were already copied during copy of columns, rows and cells, but piecemeal (e.g. multi-cell
        // sparkline group could be split into group-per-cell). Since this is a copy of whole sheet, just remove
        // the piecemeal copy and copy it propertly.
        targetSheet.SparklineGroupsInternal.RemoveAll();
        SparklineGroups.CopyTo(targetSheet);
        MergedRanges.ForEach(mr => targetSheet.Range(((XlsxRangeAddress)mr.RangeAddress).WithoutWorksheet()).Merge());
        SelectedRanges.ForEach(sr => targetSheet.SelectedRanges.Add(targetSheet.Range(((XlsxRangeAddress)sr.RangeAddress).WithoutWorksheet())));

        if (AutoFilter.IsEnabled)
        {
            var range = targetSheet.Range(((XlsxRangeAddress)AutoFilter.Range.RangeAddress).WithoutWorksheet());
            range.SetAutoFilter();
        }

        return targetSheet;
    }

    internal XlsxHyperlinks Hyperlinks { get; }

    IXlsxHyperlinks IXlsxWorksheet.Hyperlinks => Hyperlinks;

    IXlsxDataValidations IXlsxWorksheet.DataValidations
    {
        get { return DataValidations; }
    }

    private XlsxWorksheetVisibility _visibility;

    public XlsxWorksheetVisibility Visibility
    {
        get { return _visibility; }
        set
        {
            if (value != XlsxWorksheetVisibility.Visible)
                TabSelected = false;

            _visibility = value;
        }
    }

    public IXlsxWorksheet Hide()
    {
        Visibility = XlsxWorksheetVisibility.Hidden;
        return this;
    }

    public IXlsxWorksheet Unhide()
    {
        Visibility = XlsxWorksheetVisibility.Visible;
        return this;
    }

    IXlsxSheetProtection IXlsxProtectable<IXlsxSheetProtection, XlsxSheetProtectionElements>.Protection
    {
        get => Protection;
        set => Protection = (XlsxSheetProtection)value;
    }

    public IXlsxSheetProtection Protect(Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        return Protection.Protect(algorithm);
    }

    public IXlsxSheetProtection Protect(XlsxSheetProtectionElements allowedElements)
        => Protection.Protect(allowedElements);

    public IXlsxSheetProtection Protect(Algorithm algorithm, XlsxSheetProtectionElements allowedElements)
        => Protection.Protect(algorithm, allowedElements);

    public IXlsxSheetProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm)
    {
        return Protection.Protect(password, algorithm, XlsxSheetProtectionElements.SelectEverything);
    }

    public IXlsxSheetProtection Protect(String password, Algorithm algorithm, XlsxSheetProtectionElements allowedElements)
    {
        return Protection.Protect(password, algorithm, allowedElements);
    }

    IXlsxElementProtection IXlsxProtectable.Protect(Algorithm algorithm)
    {
        return Protect(algorithm);
    }

    IXlsxElementProtection IXlsxProtectable.Protect(String password, Algorithm algorithm)
    {
        return Protect(password, algorithm);
    }

    public IXlsxSheetProtection Unprotect()
    {
        return Protection.Unprotect();
    }

    public IXlsxSheetProtection Unprotect(String password)
    {
        return Protection.Unprotect(password);
    }

    IXlsxElementProtection IXlsxProtectable.Unprotect()
    {
        return Unprotect();
    }

    IXlsxElementProtection IXlsxProtectable.Unprotect(String password)
    {
        return Unprotect(password);
    }

    public new IXlsxRange Sort()
    {
        return GetRangeForSort().Sort();
    }

    public new IXlsxRange Sort(String columnsToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending,
                             Boolean matchCase = false, Boolean ignoreBlanks = true)
    {
        return GetRangeForSort().Sort(columnsToSortBy, sortOrder, matchCase, ignoreBlanks);
    }

    public new IXlsxRange Sort(Int32 columnToSortBy, XlsxSortOrder sortOrder = XlsxSortOrder.Ascending,
                             Boolean matchCase = false, Boolean ignoreBlanks = true)
    {
        return GetRangeForSort().Sort(columnToSortBy, sortOrder, matchCase, ignoreBlanks);
    }

    public new IXlsxRange SortLeftToRight(XlsxSortOrder sortOrder = XlsxSortOrder.Ascending, Boolean matchCase = false,
                                        Boolean ignoreBlanks = true)
    {
        return GetRangeForSort().SortLeftToRight(sortOrder, matchCase, ignoreBlanks);
    }

    public Boolean ShowFormulas { get; set; }

    public Boolean ShowGridLines { get; set; }

    public Boolean ShowOutlineSymbols { get; set; }

    public Boolean ShowRowColHeaders { get; set; }

    public Boolean ShowRuler { get; set; }

    public Boolean ShowWhiteSpace { get; set; }

    public Boolean ShowZeros { get; set; }

    public IXlsxWorksheet SetShowFormulas()
    {
        ShowFormulas = true;
        return this;
    }

    public IXlsxWorksheet SetShowFormulas(Boolean value)
    {
        ShowFormulas = value;
        return this;
    }

    public IXlsxWorksheet SetShowGridLines()
    {
        ShowGridLines = true;
        return this;
    }

    public IXlsxWorksheet SetShowGridLines(Boolean value)
    {
        ShowGridLines = value;
        return this;
    }

    public IXlsxWorksheet SetShowOutlineSymbols()
    {
        ShowOutlineSymbols = true;
        return this;
    }

    public IXlsxWorksheet SetShowOutlineSymbols(Boolean value)
    {
        ShowOutlineSymbols = value;
        return this;
    }

    public IXlsxWorksheet SetShowRowColHeaders()
    {
        ShowRowColHeaders = true;
        return this;
    }

    public IXlsxWorksheet SetShowRowColHeaders(Boolean value)
    {
        ShowRowColHeaders = value;
        return this;
    }

    public IXlsxWorksheet SetShowRuler()
    {
        ShowRuler = true;
        return this;
    }

    public IXlsxWorksheet SetShowRuler(Boolean value)
    {
        ShowRuler = value;
        return this;
    }

    public IXlsxWorksheet SetShowWhiteSpace()
    {
        ShowWhiteSpace = true;
        return this;
    }

    public IXlsxWorksheet SetShowWhiteSpace(Boolean value)
    {
        ShowWhiteSpace = value;
        return this;
    }

    public IXlsxWorksheet SetShowZeros()
    {
        ShowZeros = true;
        return this;
    }

    public IXlsxWorksheet SetShowZeros(Boolean value)
    {
        ShowZeros = value;
        return this;
    }

    public XlsxColor TabColor { get; set; }

    public IXlsxWorksheet SetTabColor(XlsxColor color)
    {
        TabColor = color;
        return this;
    }

    public Boolean TabSelected { get; set; }

    public Boolean TabActive
    {
        get { return _tabActive; }
        set
        {
            if (value && !_tabActive)
            {
                foreach (XlsxWorksheet ws in Worksheet.Workbook.WorksheetsInternal)
                    ws._tabActive = false;
            }
            _tabActive = value;
        }
    }

    public IXlsxWorksheet SetTabSelected()
    {
        TabSelected = true;
        return this;
    }

    public IXlsxWorksheet SetTabSelected(Boolean value)
    {
        TabSelected = value;
        return this;
    }

    public IXlsxWorksheet SetTabActive()
    {
        TabActive = true;
        return this;
    }

    public IXlsxWorksheet SetTabActive(Boolean value)
    {
        TabActive = value;
        return this;
    }

    IXlsxPivotTable IXlsxWorksheet.PivotTable(String name)
    {
        return PivotTable(name);
    }

    IXlsxPivotTables IXlsxWorksheet.PivotTables => PivotTables;

    public XlsxPivotTables PivotTables { get; }

    public Boolean RightToLeft { get; set; }

    public IXlsxWorksheet SetRightToLeft()
    {
        RightToLeft = true;
        return this;
    }

    public IXlsxWorksheet SetRightToLeft(Boolean value)
    {
        RightToLeft = value;
        return this;
    }

    public override XlsxRanges Ranges(String ranges)
    {
        var retVal = new XlsxRanges(Workbook);
        foreach (string rangeAddressStr in ranges.Split(',').Select(s => s.Trim()))
        {
            if (rangeAddressStr.StartsWith("#REF!"))
            {
                continue;
            }
            else if (XlsxHelper.IsValidRangeAddress(rangeAddressStr))
            {
                retVal.Add(Range(new XlsxRangeAddress(Worksheet, rangeAddressStr)));
            }
            else if (DefinedNames.TryGetValue(rangeAddressStr, out var worksheetNamedRange))
            {
                worksheetNamedRange.Ranges.ForEach(retVal.Add);
            }
            else if (Workbook.DefinedNames.TryGetValue(rangeAddressStr, out var workbookDefinedName)
                && workbookDefinedName.Ranges.First().Worksheet == this)
            {
                workbookDefinedName.Ranges.ForEach(retVal.Add);
            }
        }
        return retVal;
    }

    IXlsxAutoFilter IXlsxWorksheet.AutoFilter
    {
        get { return AutoFilter; }
    }

    public IXlsxRows RowsUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents, Func<IXlsxRow, Boolean>? predicate = null)
    {
        var rows = new XlsxRows(Workbook, worksheet: null, this);
        var rowsUsed = new HashSet<Int32>();
        foreach (var rowNum in Internals.RowsCollection.Keys.Concat(Internals.CellsCollection.RowsUsedKeys))
        {
            if (!rowsUsed.Add(rowNum))
            {
                continue;
            }
            var row = Row(rowNum);
            if (!row.IsEmpty(options) && (predicate == null || predicate(row)))
                rows.Add(row);
        }
        return rows;
    }

    public IXlsxRows RowsUsed(Func<IXlsxRow, Boolean>? predicate = null)
    {
        return RowsUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    public IXlsxColumns ColumnsUsed(XlsxCellsUsedOptions options = XlsxCellsUsedOptions.AllContents, Func<IXlsxColumn, Boolean>? predicate = null)
    {
        var columns = new XlsxColumns(Workbook, worksheet: null, defaultStyleSheet: this);
        var columnsUsed = new HashSet<Int32>();
        Internals.ColumnsCollection.Keys.ForEach(r => columnsUsed.Add(r));
        Internals.CellsCollection.ColumnsUsedKeys.ForEach(r => columnsUsed.Add(r));
        foreach (var columnNum in columnsUsed)
        {
            var column = Column(columnNum);
            if (!column.IsEmpty(options) && (predicate == null || predicate(column)))
                columns.Add(column);
        }
        return columns;
    }

    public IXlsxColumns ColumnsUsed(Func<IXlsxColumn, Boolean>? predicate = null)
    {
        return ColumnsUsed(XlsxCellsUsedOptions.AllContents, predicate);
    }

    internal void RegisterRangeIndex(IXlsxRangeIndex rangeIndex)
    {
        _rangeIndices.Add(rangeIndex);
    }

    internal void Cleanup()
    {
        Internals.Dispose();
        foreach (var picture in Pictures)
            picture.Dispose();

        _rangeRepository.Clear();
        _rangeIndices.Clear();
    }

    #endregion IXlsxWorksheet Members

    #region Outlines

    public void IncrementColumnOutline(Int32 level)
    {
        if (level <= 0) return;
        if (_columnOutlineCount.TryGetValue(level, out Int32 value))
            _columnOutlineCount[level] = value + 1;
        else
            _columnOutlineCount.Add(level, 1);
    }

    public void DecrementColumnOutline(Int32 level)
    {
        if (level <= 0) return;
        if (_columnOutlineCount.TryGetValue(level, out Int32 value))
        {
            if (value > 0)
                _columnOutlineCount[level] = value - 1;
        }
        else
            _columnOutlineCount.Add(level, 0);
    }

    public Int32 GetMaxColumnOutline()
    {
        var list = _columnOutlineCount.Where(kp => kp.Value > 0).ToList();
        return list.Count == 0 ? 0 : list.Max(kp => kp.Key);
    }

    public void IncrementRowOutline(Int32 level)
    {
        if (level <= 0) return;
        if (_rowOutlineCount.TryGetValue(level, out Int32 value))
            _rowOutlineCount[level] = value + 1;
        else
            _rowOutlineCount.Add(level, 0);
    }

    public void DecrementRowOutline(Int32 level)
    {
        if (level <= 0) return;
        if (_rowOutlineCount.TryGetValue(level, out Int32 value))
        {
            if (value > 0)
                _rowOutlineCount[level] = level - 1;
        }
        else
            _rowOutlineCount.Add(level, 0);
    }

    public Int32 GetMaxRowOutline()
    {
        return _rowOutlineCount.Count == 0 ? 0 : _rowOutlineCount.Where(kp => kp.Value > 0).Max(kp => kp.Key);
    }

    #endregion Outlines

    public XlsxRow? FirstRowUsed()
    {
        return FirstRowUsed(XlsxCellsUsedOptions.AllContents);
    }

    public XlsxRow? FirstRowUsed(XlsxCellsUsedOptions options)
    {
        var rngRow = AsRange().FirstRowUsed(options);
        return rngRow != null ? Row(rngRow.RangeAddress.FirstAddress.RowNumber) : null;
    }

    public XlsxRow? LastRowUsed()
    {
        return LastRowUsed(XlsxCellsUsedOptions.AllContents);
    }

    public XlsxRow? LastRowUsed(XlsxCellsUsedOptions options)
    {
        var rngRow = AsRange().LastRowUsed(options);
        return rngRow != null ? Row(rngRow.RangeAddress.LastAddress.RowNumber) : null;
    }

    public XlsxColumn LastColumn()
    {
        return Column(XlsxHelper.MaxColumnNumber);
    }

    public XlsxColumn FirstColumn()
    {
        return Column(1);
    }

    public XlsxRow FirstRow()
    {
        return Row(1);
    }

    public XlsxRow LastRow()
    {
        return Row(XlsxHelper.MaxRowNumber);
    }

    public XlsxColumn? FirstColumnUsed()
    {
        return FirstColumnUsed(XlsxCellsUsedOptions.AllContents);
    }

    public XlsxColumn? FirstColumnUsed(XlsxCellsUsedOptions options)
    {
        var rngColumn = AsRange().FirstColumnUsed(options);
        return rngColumn != null ? Column(rngColumn.RangeAddress.FirstAddress.ColumnNumber) : null;
    }

    public XlsxColumn? LastColumnUsed()
    {
        return LastColumnUsed(XlsxCellsUsedOptions.AllContents);
    }

    public XlsxColumn? LastColumnUsed(XlsxCellsUsedOptions options)
    {
        var rngColumn = AsRange().LastColumnUsed(options);
        return rngColumn != null ? Column(rngColumn.RangeAddress.LastAddress.ColumnNumber) : null;
    }

    public XlsxRow Row(Int32 row)
    {
        return Row(row, true);
    }

    public XlsxColumn Column(Int32 columnNumber)
    {
        if (columnNumber <= 0 || columnNumber > XlsxHelper.MaxColumnNumber)
            throw new ArgumentOutOfRangeException(nameof(columnNumber), $"Column number must be between 1 and {XlsxHelper.MaxColumnNumber}");

        if (Internals.ColumnsCollection.TryGetValue(columnNumber, out XlsxColumn column))
            return column;
        else
        {
            // This is a new column so we're going to reference all
            // cells in this column to preserve their formatting
            Internals.RowsCollection.Keys.ForEach(r => Cell(r, columnNumber).PingStyle());

            column = RangeFactory.CreateColumn(columnNumber);
            Internals.ColumnsCollection.Add(columnNumber, column);
        }

        return column;
    }

    public IXlsxColumn Column(String column)
    {
        return Column(XlsxHelper.GetColumnNumberFromLetter(column));
    }

    public override XlsxRange AsRange()
    {
        return Range(1, 1, XlsxHelper.MaxRowNumber, XlsxHelper.MaxColumnNumber);
    }

    internal override void WorksheetRangeShiftedColumns(XlsxRange range, int columnsShifted)
    {
        if (!range.IsEntireColumn())
        {
            var model = new XlsxRangeAddress(
                range.RangeAddress.FirstAddress,
                new XlsxAddress(range.RangeAddress.LastAddress.RowNumber, XlsxHelper.MaxColumnNumber, false, false));
            var rangesToSplit = Worksheet.MergedRanges
                .GetIntersectedRanges(model)
                .Where(r => r.RangeAddress.FirstAddress.RowNumber < range.RangeAddress.FirstAddress.RowNumber ||
                            r.RangeAddress.LastAddress.RowNumber > range.RangeAddress.LastAddress.RowNumber)
                .ToList();
            foreach (var rangeToSplit in rangesToSplit)
            {
                Worksheet.MergedRanges.Remove(rangeToSplit);
            }
        }

        ShiftPageBreaksColumns(range, columnsShifted);

        var sheetListeners = GetSheetListeners();

        if (columnsShifted > 0)
        {
            var insertedArea = Aid.Xlsx.Model.Grid.Coordinates.Area
                .FromRangeAddress(range.RangeAddress)
                .SliceFromLeft(1)
                .ExtendRight(columnsShifted - 1);
            foreach (var listener in sheetListeners)
                listener.OnInsertAreaAndShiftRight(range.Worksheet, insertedArea);
        }
        else if (columnsShifted < 0)
        {
            var area = Aid.Xlsx.Model.Grid.Coordinates.Area.FromRangeAddress(range.RangeAddress);
            foreach (var listener in sheetListeners)
                listener.OnDeleteAreaAndShiftLeft(range.Worksheet, area);
        }
    }

    private void ShiftPageBreaksColumns(XlsxRange range, int columnsShifted)
    {
        for (var i = 0; i < PageSetup.ColumnBreaks.Count; i++)
        {
            int br = PageSetup.ColumnBreaks[i];
            if (range.RangeAddress.FirstAddress.ColumnNumber <= br)
            {
                PageSetup.ColumnBreaks[i] = br + columnsShifted;
            }
        }
    }

    internal override void WorksheetRangeShiftedRows(XlsxRange range, int rowsShifted)
    {
        if (!range.IsEntireRow())
        {
            var model = new XlsxRangeAddress(
                range.RangeAddress.FirstAddress,
                new XlsxAddress(XlsxHelper.MaxRowNumber, range.RangeAddress.LastAddress.ColumnNumber, false, false));
            var rangesToSplit = Worksheet.MergedRanges
                .GetIntersectedRanges(model)
                .Where(r => r.RangeAddress.FirstAddress.ColumnNumber < range.RangeAddress.FirstAddress.ColumnNumber ||
                            r.RangeAddress.LastAddress.ColumnNumber > range.RangeAddress.LastAddress.ColumnNumber)
                .ToList();
            foreach (var rangeToSplit in rangesToSplit)
            {
                Worksheet.MergedRanges.Remove(rangeToSplit);
            }
        }

        ShiftPageBreaksRows(range, rowsShifted);

        var sheetListeners = GetSheetListeners();

        if (rowsShifted > 0)
        {
            var insertedArea = Aid.Xlsx.Model.Grid.Coordinates.Area
                .FromRangeAddress(range.RangeAddress)
                .SliceFromTop(1)
                .ExtendBelow(rowsShifted - 1);
            foreach (var listener in sheetListeners)
                listener.OnInsertAreaAndShiftDown(range.Worksheet, insertedArea);
        }
        else if (rowsShifted < 0)
        {
            var area = Aid.Xlsx.Model.Grid.Coordinates.Area.FromRangeAddress(range.RangeAddress);
            foreach (var listener in sheetListeners)
                listener.OnDeleteAreaAndShiftUp(range.Worksheet, area);
        }
    }

    private List<ISheetListener> GetSheetListeners()
    {
        var sheetListeners = new List<ISheetListener>
        {
            Workbook.CalcEngine,
            Hyperlinks,
            Workbook.DefinedNamesInternal,
            DataValidations,
        };
        foreach (var worksheet in Workbook.WorksheetsInternal)
            sheetListeners.Add(worksheet.DefinedNames);

        sheetListeners.AddRange(SparklineGroupsInternal);

        // CF can contain formulas for any worksheet, notify about all changes
        foreach (var worksheet in Workbook.WorksheetsInternal)
            sheetListeners.Add(worksheet.ConditionalFormats);

        return sheetListeners;
    }

    private void ShiftPageBreaksRows(XlsxRange range, int rowsShifted)
    {
        for (var i = 0; i < PageSetup.RowBreaks.Count; i++)
        {
            int br = PageSetup.RowBreaks[i];
            if (range.RangeAddress.FirstAddress.RowNumber <= br)
            {
                PageSetup.RowBreaks[i] = br + rowsShifted;
            }
        }
    }

    public void NotifyRangeShiftedRows(XlsxRange range, Int32 rowsShifted)
    {
        var rangesToShift = _rangeRepository
            .Where(r => r.RangeAddress.IsValid)
            .OrderBy(r => r.RangeAddress.FirstAddress.RowNumber * -Math.Sign(rowsShifted))
            .ToList();

        WorksheetRangeShiftedRows(range, rowsShifted);

        bool collapsed = false;
        foreach (var storedRange in rangesToShift)
        {
            if (storedRange.IsEntireColumn())
                continue;

            if (ReferenceEquals(range, storedRange))
                continue;

            storedRange.WorksheetRangeShiftedRows(range, rowsShifted);
            if (range.RangeAddress == storedRange.RangeAddress)
            {
                collapsed = true;
            }
        }
        if (!collapsed)
        {
            range.WorksheetRangeShiftedRows(range, rowsShifted);
        }
    }

    public void NotifyRangeShiftedColumns(XlsxRange range, Int32 columnsShifted)
    {
        var rangesToShift = _rangeRepository
            .Where(r => r.RangeAddress.IsValid)
            .OrderBy(r => r.RangeAddress.FirstAddress.ColumnNumber * -Math.Sign(columnsShifted))
            .ToList();

        WorksheetRangeShiftedColumns(range, columnsShifted);

        bool collapsed = false;
        foreach (var storedRange in rangesToShift)
        {
            if (storedRange.IsEntireRow())
                continue;

            if (ReferenceEquals(range, storedRange))
                continue;

            storedRange.WorksheetRangeShiftedColumns(range, columnsShifted);
            if (range.RangeAddress == storedRange.RangeAddress)
            {
                collapsed = true;
            }
        }
        if (!collapsed)
        {
            range.WorksheetRangeShiftedColumns(range, columnsShifted);
        }
    }

    internal XlsxRow Row(Int32 rowNumber, Boolean pingCells)
    {
        if (rowNumber <= 0 || rowNumber > XlsxHelper.MaxRowNumber)
            throw new ArgumentOutOfRangeException(nameof(rowNumber), $"Row number must be between 1 and {XlsxHelper.MaxRowNumber}");

        if (Internals.RowsCollection.TryGetValue(rowNumber, out XlsxRow row))
            return row;
        else
        {
            if (pingCells)
            {
                // This is a new row so we're going to reference all
                // cells in columns of this row to preserve their formatting
                Internals.ColumnsCollection.Keys.ForEach(c => Cell(rowNumber, c).PingStyle());
            }

            row = RangeFactory.CreateRow(rowNumber);
            Internals.RowsCollection.Add(rowNumber, row);
        }

        return row;
    }

    public IXlsxTable Table(XlsxRange range, Boolean addToTables, Boolean setAutofilter = true)
    {
        return Table(range, TableNameGenerator.GetNewTableName(Workbook), addToTables, setAutofilter);
    }

    public IXlsxTable Table(XlsxRange range, String name, Boolean addToTables, Boolean setAutofilter = true)
    {
        CheckRangeNotOverlappingOtherEntities(range);
        XlsxRangeAddress rangeAddress;
        if (range.Rows().Count() == 1)
        {
            rangeAddress = new XlsxRangeAddress(range.FirstCell().Address, range.LastCell().CellBelow().Address);
            range.InsertRowsBelow(1);
        }
        else
            rangeAddress = range.RangeAddress;

        var rangeKey = new XlsxRangeKey(XlsxRangeType.Table, rangeAddress);
        var table = (XlsxTable)_rangeRepository.GetOrCreate(ref rangeKey);

        if (table.Name != name)
            table.Name = name;

        if (addToTables && !Tables.Contains(table))
        {
            Tables.Add(table);
        }

        if (setAutofilter && !table.ShowAutoFilter)
            table.InitializeAutoFilter();

        return table;
    }

    private void CheckRangeNotOverlappingOtherEntities(XlsxRange range)
    {
        // Check that the range doesn't overlap with any existing tables
        var firstOverlappingTable = Tables.FirstOrDefault<XlsxTable>(t => t.RangeUsed().Intersects(range));
        if (firstOverlappingTable != null)
            throw new InvalidOperationException($"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} is already part of table '{firstOverlappingTable.Name}'");

        // Check that the range doesn't overlap with any filters
        if (AutoFilter.IsEnabled && this.AutoFilter.Range.Intersects(range))
            throw new InvalidOperationException($"The range {range.RangeAddress.ToStringRelative(includeSheet: true)} overlaps with the worksheet's autofilter.");
    }

    private IXlsxRange GetRangeForSort()
    {
        var range = RangeUsed();
        SortColumns.ForEach(e => range.SortColumns.Add(e.ElementNumber, e.SortOrder, e.IgnoreBlanks, e.MatchCase));
        SortRows.ForEach(e => range.SortRows.Add(e.ElementNumber, e.SortOrder, e.IgnoreBlanks, e.MatchCase));
        return range;
    }

    public XlsxPivotTable PivotTable(String name)
    {
        return (XlsxPivotTable)PivotTables.PivotTable(name);
    }

    public override IXlsxCells Cells()
    {
        return Cells(true, XlsxCellsUsedOptions.All);
    }

    public override XlsxCells Cells(Boolean usedCellsOnly)
    {
        if (usedCellsOnly)
            return Cells(true, XlsxCellsUsedOptions.AllContents);
        else
            return Range((this as IXlsxRangeBase).FirstCellUsed(XlsxCellsUsedOptions.All),
                         (this as IXlsxRangeBase).LastCellUsed(XlsxCellsUsedOptions.All))
                .Cells(false, XlsxCellsUsedOptions.All);
    }

    public override XlsxCell? Cell(String cellAddressInRange)
    {
        var cell = base.Cell(cellAddressInRange);
        if (cell is not null)
            return cell;

        if (Workbook.DefinedNames.TryGetValue(cellAddressInRange, out var definedName))
        {
            if (definedName.Ranges.Count == 0)
                return null;

            return definedName.Ranges.First().FirstCell().CastTo<XlsxCell>();
        }

        return null;
    }

    public override XlsxRange? Range(String rangeAddressStr)
    {
        if (XlsxHelper.IsValidRangeAddress(rangeAddressStr))
            return Range(new XlsxRangeAddress(Worksheet, rangeAddressStr));

        if (rangeAddressStr.Contains('['))
            return Table(rangeAddressStr[..rangeAddressStr.IndexOf("[")]) as XlsxRange;

        if (DefinedNames.TryGetValue(rangeAddressStr, out var sheetDefinedName))
            return sheetDefinedName.Ranges.First().CastTo<XlsxRange>();

        if (Workbook.DefinedNamesInternal.TryGetValue(rangeAddressStr, out var workbookDefinedName))
        {
            if (workbookDefinedName.Ranges.Count == 0)
                return null;

            return workbookDefinedName.Ranges.First().CastTo<XlsxRange>();
        }

        return null;
    }

    public IXlsxRanges MergedRanges { get { return Internals.MergedRanges; } }

    IXlsxConditionalFormats IXlsxWorksheet.ConditionalFormats => ConditionalFormats;

    internal XlsxConditionalFormats ConditionalFormats { get; }

    public IXlsxSparklineGroups SparklineGroups => SparklineGroupsInternal;

    public IXlsxRanges SelectedRanges
    {
        get
        {
            _selectedRanges.RemoveAll(r => !r.RangeAddress.IsValid);
            return _selectedRanges;
        }
    }

    IXlsxCell? IXlsxWorksheet.ActiveCell
    {
        get => ActiveCell is not null ? new XlsxCell(this, ActiveCell.Value) : null;
        set => ActiveCell = value is not null ? Point.FromAddress(value.Address) : null;
    }

    /// <summary>
    /// Address of active cell/cursor in the worksheet.
    /// </summary>
    internal Point? ActiveCell { get; set; }

    private XlsxCalcEngine CalcEngine => Workbook.CalcEngine;

    public XlsxCellValue Evaluate(String expression, string? formulaAddress = null)
    {
        IXlsxAddress? address = formulaAddress is not null ? XlsxAddress.Create(formulaAddress) : null;
        return CalcEngine.EvaluateFormula(expression, Workbook, this, address, true).ToCellValue();
    }

    public void RecalculateAllFormulas()
    {
        Internals.CellsCollection.FormulaSlice.MarkDirty(Aid.Xlsx.Model.Grid.Coordinates.Area.Full);
        Workbook.CalcEngine.Recalculate(Workbook, Name);
    }

    public String Author { get; set; }

    public override string ToString()
    {
        return this.Name;
    }

    IXlsxPictures IXlsxWorksheet.Pictures => Pictures;

    public Boolean IsPasswordProtected => Protection.IsPasswordProtected;

    public bool IsProtected => Protection.IsProtected;

    internal XlsxPictures Pictures { get; }

    public IXlsxPicture Picture(string pictureName)
    {
        return Pictures.Picture(pictureName);
    }

    public IXlsxPicture AddPicture(Stream stream)
    {
        return Pictures.Add(stream);
    }

    public IXlsxPicture AddPicture(Stream stream, string name)
    {
        return Pictures.Add(stream, name);
    }

    internal IXlsxPicture AddPicture(Stream stream, string name, int Id)
    {
        return ((XlsxPictures)Pictures).Add(stream, name, Id);
    }

    public IXlsxPicture AddPicture(Stream stream, XlsxPictureFormat format)
    {
        return Pictures.Add(stream, format);
    }

    public IXlsxPicture AddPicture(Stream stream, XlsxPictureFormat format, string name)
    {
        return Pictures.Add(stream, format, name);
    }

    public IXlsxPicture AddPicture(string imageFile)
    {
        return Pictures.Add(imageFile);
    }

    public IXlsxPicture AddPicture(string imageFile, string name)
    {
        return Pictures.Add(imageFile, name);
    }

    public override Boolean IsEntireRow()
    {
        return true;
    }

    public override Boolean IsEntireColumn()
    {
        return true;
    }

    internal IXlsxTable InsertTable(Point origin, IInsertDataReader reader, String? tableName, Boolean createTable, Boolean addHeadings, Boolean transpose)
    {
        if (createTable && Tables.Any<XlsxTable>(t => t.Area.Contains(origin)))
            throw new InvalidOperationException($"This cell '{origin}' is already part of a table.");

        var range = InsertData(origin, reader, addHeadings, transpose);

        if (createTable)
            // Create a table and save it in the file
            return tableName == null ? range.CreateTable() : range.CreateTable(tableName);
        else
            // Create a table, but keep it in memory. Saved file will contain only "raw" data and column headers
            return tableName == null ? range.AsTable() : range.AsTable(tableName);
    }

    internal XlsxRange InsertData(Point origin, IInsertDataReader reader, Boolean addHeadings, Boolean transpose)
    {
        // Prepare data. Heading is basically just another row of data, so unify it.
        var rows = reader.GetRecords();
        var propCount = reader.GetPropertiesCount();
        if (addHeadings)
        {
            var headings = new XlsxCellValue[propCount];
            for (var i = 0; i < propCount; i++)
                headings[i] = reader.GetPropertyName(i);

            rows = new[] { headings }.Concat(rows);
        }

        if (transpose)
        {
            rows = TransposeJaggedArray(rows);
        }

        var valueSlice = Internals.CellsCollection.ValueSlice;
        var formatSlice = Internals.CellsCollection.FormatSlice;

        // A buffer to avoid multiple enumerations of the source.
        var rowBuffer = new List<XlsxCellValue>();
        var maximumColumn = origin.Column;
        var rowNumber = origin.Row;
        foreach (var row in rows)
        {
            rowBuffer.AddRange(row);

            // InsertData should also clear data and if row doesn't have enough data,
            // fill in the rest. Only fill up to the props to be consistent. We can't
            // know how long any next row will be, so props are used as a source of truth
            // for which columns should be cleared.
            for (var i = rowBuffer.Count; i < propCount; ++i)
                rowBuffer.Add(Blank.Value);

            // Each row can have different number of values, so we have to check every row.
            maximumColumn = Math.Max(origin.Column + rowBuffer.Count - 1, maximumColumn);
            if (maximumColumn > XlsxHelper.MaxColumnNumber || rowNumber > XlsxHelper.MaxRowNumber)
                throw new ArgumentException("Data would write out of the sheet.");

            var column = origin.Column;
            for (var i = 0; i < rowBuffer.Count; ++i)
            {
                var value = rowBuffer[i];
                var point = new Point(rowNumber, column);
                var modifiedStyle = GetStyleForValue(value, point);
                if (modifiedStyle is not null)
                {
                    if (value.IsText && value.GetText()[0] == '\'')
                        value = value.GetText()[1..];

                    formatSlice.Set(point, modifiedStyle);
                }

                valueSlice.SetCellValue(point, value);
                if ((!addHeadings || rowNumber != origin.Row) && !transpose && i < propCount
                    && reader.GetNumberFormat(i) is { } numberFormat)
                {
                    Cell(point.Row, point.Column).Style.NumberFormat.Format = numberFormat;
                }
                column++;
            }

            rowBuffer.Clear();
            rowNumber++;
        }

        // If there is no row, rowNumber is kept at origin instead of last row + 1 .
        var lastRow = Math.Max(rowNumber - 1, origin.Row);
        var insertedArea = new Area(origin, new Point(lastRow, maximumColumn));

        // If inserted area affected a table, we must fix headings and totals, because these values
        // are duplicated. Basically the table values are the truth and cells are a reflection of the
        // truth, but here we inserted shadow first.
        foreach (var table in Tables)
            table.RefreshFieldsFromCells(insertedArea);

        // Invalidate only once, not for every cell.
        Workbook.CalcEngine.MarkDirty(Worksheet, insertedArea);

        // Return area that contains all inserted cells, no matter how jagged were data.
        return Range(
            insertedArea.FirstPoint.Row,
            insertedArea.FirstPoint.Column,
            insertedArea.LastPoint.Row,
            insertedArea.LastPoint.Column);

        // Rather memory inefficient, but the original code also materialized
        // data through Linq/required multiple enumerations.
        static List<List<XlsxCellValue>> TransposeJaggedArray(IEnumerable<IEnumerable<XlsxCellValue>> enumerable)
        {
            var destination = new List<List<XlsxCellValue>>();

            var sourceRow = 1;
            foreach (var row in enumerable)
            {
                var sourceColumn = 1;
                foreach (var sourceValue in row)
                {
                    // The original has `sourceValue` at [sourceRow, sourceColumn]
                    var destinationRowCount = destination.Count;
                    if (sourceColumn > destinationRowCount)
                        destination.Add([]);

                    // There can be jagged arrays and the destination can have spaces between columns.
                    var destinationRow = destination[sourceColumn - 1];
                    while (destinationRow.Count < sourceRow - 1)
                        destinationRow.Add(Blank.Value);

                    destinationRow.Add(sourceValue);
                    sourceColumn++;
                }

                sourceRow++;
            }

            return destination;
        }
    }

    /// <summary>
    /// Get cell or null, if cell doesn't exist.
    /// </summary>
    internal XlsxCell? GetCell(Point point)
    {
        return Worksheet.Internals.CellsCollection.GetUsedCell(point);
    }

    public XlsxRange GetOrCreateRange(XlsxRangeAddress rangeAddress)
    {
        var rangeKey = new XlsxRangeKey(XlsxRangeType.Range, rangeAddress);
        var range = _rangeRepository.GetOrCreate(ref rangeKey);

        return (XlsxRange)range;
    }

    /// <summary>
    /// Get a range row from the shared repository or create a new one.
    /// </summary>
    /// <param name="address">Address of range row.</param>
    /// <returns>Range row with the specified address.</returns>
    public XlsxRangeRow RangeRow(XlsxRangeAddress address)
    {
        var rangeKey = new XlsxRangeKey(XlsxRangeType.RangeRow, address);
        var rangeRow = (XlsxRangeRow)_rangeRepository.GetOrCreate(ref rangeKey);
        return rangeRow;
    }

    /// <summary>
    /// Get a range column from the shared repository or create a new one.
    /// </summary>
    /// <param name="address">Address of range column.</param>
    /// <returns>Range column with the specified address.</returns>
    public XlsxRangeColumn RangeColumn(XlsxRangeAddress address)
    {
        var rangeKey = new XlsxRangeKey(XlsxRangeType.RangeColumn, address);
        var rangeColumn = (XlsxRangeColumn)_rangeRepository.GetOrCreate(ref rangeKey);
        return rangeColumn;
    }

    protected override void OnRangeAddressChanged(XlsxRangeAddress oldAddress, XlsxRangeAddress newAddress)
    {
    }

    public void RelocateRange(XlsxRangeType rangeType, XlsxRangeAddress oldAddress, XlsxRangeAddress newAddress)
    {
        var oldKey = new XlsxRangeKey(rangeType, oldAddress);
        var newKey = new XlsxRangeKey(rangeType, newAddress);
        var range = _rangeRepository.Replace(ref oldKey, ref newKey);

        foreach (var rangeIndex in _rangeIndices)
        {
            if (!rangeIndex.MatchesType(rangeType))
                continue;

            if (rangeIndex.Remove(oldAddress) &&
                newAddress.IsValid &&
                range != null)
            {
                rangeIndex.Add(range);
            }
        }
    }

    internal void DeleteColumn(int columnNumber)
    {
        Internals.ColumnsCollection.Remove(columnNumber);

        var columnsToMove = new List<Int32>(Internals.ColumnsCollection.Where(c => c.Key > columnNumber).Select(c => c.Key).OrderBy(c => c));
        foreach (int column in columnsToMove)
        {
            Internals.ColumnsCollection.Add(column - 1, Internals.ColumnsCollection[column]);
            Internals.ColumnsCollection.Remove(column);

            Internals.ColumnsCollection[column - 1].SetColumnNumber(column - 1);
        }
    }

    internal void DeleteRow(int rowNumber)
    {
        Internals.RowsCollection.Remove(rowNumber);

        var rowsToMove = new List<Int32>(Internals.RowsCollection.Where(c => c.Key > rowNumber).Select(c => c.Key).OrderBy(r => r));
        foreach (int row in rowsToMove)
        {
            Internals.RowsCollection.Add(row - 1, Worksheet.Internals.RowsCollection[row]);
            Internals.RowsCollection.Remove(row);

            Internals.RowsCollection[row - 1].SetRowNumber(row - 1);
        }
    }

    internal void DeleteRange(XlsxRangeAddress rangeAddress)
    {
        var rangeKey = new XlsxRangeKey(XlsxRangeType.Range, rangeAddress);
        _rangeRepository.Remove(ref rangeKey);
    }

    /// <summary>
    /// Get the actual style for a point in the sheet.
    /// </summary>
    internal XlsxCellFormatValue GetStyleValue(Point point)
    {
        // TODO Styles: This is basically a duplication of Hierarchy.Resolve(). Investigate deduplication.
        var cellFormat = Internals.CellsCollection.FormatSlice.GetFormat(point);
        if (cellFormat is not null)
            return cellFormat;

        // TODO Styles: Ensure all cross points are set at this time (=load+change). Taking from row/col should only be done if no cross is there.
        // If the slice doesn't contain any value, determine values by inheriting.
        // Cells that lie on an intersection of a XlsxColumn and a XlsxRow have their
        // style set when column/row is created to avoid problems with correct which
        // style has precedence. I.e. set column blue, set row red => cell is red.
        // Swap order the the cell is blue.
        return GetInheritedFormat(point);
    }

    internal XlsxCellFormatValue GetInheritedFormat(Point point)
    {
        if (Internals.RowsCollection.TryGetValue(point.Row, out var row) && row.FormatValue is not null)
            return row.FormatValue;

        if (Internals.ColumnsCollection.TryGetValue(point.Column, out var column) && column.FormatValue is not null)
            return column.FormatValue;

        var sheetFormat = FormatValue;
        if (sheetFormat is not null)
            return sheetFormat;

        return Workbook.Styles.DefaultFormat;
    }

    /// <summary>
    /// Get a style that should be used for a <paramref name="value"/>,
    /// if the value is set to the <paramref name="point"/>.
    /// </summary>
    internal XlsxCellFormatValue? GetStyleForValue(XlsxCellValue value, Point point)
    {
        // Because StyleValue property retrieves value from a slice,
        // access it only if necessary. This happens during ever cell
        // of modification and thus is performance critical.
        switch (value.Type)
        {
            case XlsxDataType.DateTime:
                {
                    var onlyDatePart = value.GetUnifiedNumber() % 1 == 0;
                    var currentFormat = GetStyleValue(point);
                    if (currentFormat.NumberFormat.IsGeneralFormat())
                    {
                        var numberFormatId = onlyDatePart
                            ? XlsxPredefinedFormat.DateTime.DayMonthYear4WithSlashes
                            : XlsxPredefinedFormat.DateTime.MonthDayYear4WithDashesHour24Minutes;
                        var dateTimeNumberFormat = XlsxPredefinedFormat.FormatCodes[(int)numberFormatId];
                        return Workbook.Styles.GetModifiedFormat(currentFormat, dateTimeNumberFormat);
                    }
                }
                break;

            case XlsxDataType.TimeSpan:
                {
                    var currentFormat = GetStyleValue(point);
                    if (currentFormat.NumberFormat.IsGeneralFormat())
                    {
                        var durationNumberFormat = XlsxPredefinedFormat.FormatCodes[(int)XlsxPredefinedFormat.DateTime.Hour12MinutesSeconds];
                        return Workbook.Styles.GetModifiedFormat(currentFormat, durationNumberFormat);
                    }
                }
                break;

            case XlsxDataType.Text:
                {
                    var text = value.GetText();
                    var startsWithQuote = text.Length > 0 && text[0] == '\'';
                    var containsNewLine = text.Contains(Environment.NewLine, StringComparison.Ordinal);
                    if (!startsWithQuote && !containsNewLine)
                        break;

                    var currentFormat = GetStyleValue(point);
                    if (startsWithQuote && !currentFormat.IncludeQuotePrefix)
                    {
                        currentFormat = Workbook.Styles.GetRegisteredCellFormat(currentFormat, format => format with { IncludeQuotePrefix = true });
                    }

                    if (containsNewLine && !currentFormat.Alignment.WrapText)
                    {
                        currentFormat = Workbook.Styles.GetModifiedFormat(currentFormat, alignment => alignment with { WrapText = true });
                    }

                    return currentFormat;
                }
        }

        return null;
    }
}
