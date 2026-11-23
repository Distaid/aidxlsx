#nullable disable

using System.Diagnostics;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.PivotTables.PivotStyleFormats;
using Aid.Xlsx.Model.PivotTables.PivotValues;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PivotTables;

[DebuggerDisplay("{Name}")]
internal class XlsxPivotTable : IXlsxPivotTable
{
    private readonly XlsxWorksheet _worksheet;
    private string _name;

    /// <summary>
    /// List of all fields in the pivot table, roughly represents <c>pivotTableDefinition.
    /// pivotFields</c>. Contains info about each field, mostly page/axis info (data field can
    /// reference same field multiple times, so it mostly stores data in data fields).
    /// </summary>
    private readonly List<XlsxPivotTableField> _fields = [];
    private readonly List<XlsxPivotFormat> _formats = [];
    private readonly List<XlsxPivotConditionalFormat> _conditionalFormats = [];
    private XlsxPivotCache _cache;
    private int _filterFieldsPageWrap;
    private bool _outline = true;
    private bool _outlineData = false;
    private bool _compact = true;
    private bool _compactData = true;

    internal XlsxPivotTable(XlsxWorksheet worksheet, XlsxPivotCache cache)
    {
        _worksheet = worksheet;
        Filters = new XlsxPivotTableFilters(this);
        RowAxis = new XlsxPivotTableAxis(this, XlsxPivotAxis.AxisRow);
        ColumnAxis = new XlsxPivotTableAxis(this, XlsxPivotAxis.AxisCol);
        DataFields = new XlsxPivotDataFields(this);
        Theme = XlsxPivotTableTheme.PivotStyleLight16;
        _cache = cache;

        SetExcelDefaults();
    }

    IXlsxPivotCache IXlsxPivotTable.PivotCache { get => PivotCache; set => PivotCache = (XlsxPivotCache)value; }

    public IXlsxCell TargetCell
    {
        get
        {
            var filterRows = Filters.GetSizeWithGap().Height;
            Point tableCorner = Area.FirstPoint;
            Point targetPoint = tableCorner.ShiftRow(-filterRows);
            return _worksheet.Internals.CellsCollection.GetCell(targetPoint);
        }
        set
        {
            var filterRows = Filters.GetSizeWithGap().Height;
            Point valuePoint = ((XlsxCell)value).Point;
            Point tableCorner = valuePoint.ShiftRow(filterRows);
            Area = Area.At(tableCorner);
        }
    }

    public XlsxPivotCache PivotCache
    {
        get => _cache;
        set
        {
            IReadOnlyList<string> oldNames = _cache.FieldNames;
            _cache = value;
            UpdateCacheFields(oldNames);
        }
    }

    public IXlsxPivotFields ReportFilters => Filters;

    public IXlsxPivotFields ColumnLabels => ColumnAxis;

    public IXlsxPivotFields RowLabels => RowAxis;

    public IXlsxPivotValues Values => DataFields;

    public IEnumerable<IXlsxPivotField> ImplementedFields
    {
        get
        {
            foreach (IXlsxPivotField pf in ReportFilters)
                yield return pf;

            foreach (IXlsxPivotField pf in RowLabels)
                yield return pf;

            foreach (IXlsxPivotField pf in ColumnLabels)
                yield return pf;
        }
    }

    /// <summary>
    /// Table theme this pivot table will use.
    /// </summary>
    public XlsxPivotTableTheme Theme { get; set; }

    /// <summary>
    /// All fields reflected in the pivot cache.
    /// Order of fields is same as for in the <see cref="PivotCache"/>.
    /// </summary>
    internal IReadOnlyList<XlsxPivotTableField> PivotFields => _fields;

    internal XlsxPivotTableFilters Filters { get; }

    internal XlsxPivotTableAxis RowAxis { get; }

    internal XlsxPivotTableAxis ColumnAxis { get; }

    internal XlsxPivotDataFields DataFields { get; }

    internal IReadOnlyList<XlsxPivotFormat> Formats => _formats;

    internal IReadOnlyList<XlsxPivotConditionalFormat> ConditionalFormats => _conditionalFormats;

    public IXlsxPivotTable CopyTo(IXlsxCell targetCell)
    {
        var addressComparer = new XlsxAddressComparer(ignoreFixed: true);
        if (addressComparer.Equals(targetCell.Address, this.TargetCell.Address))
            throw new InvalidOperationException("Cannot copy pivot table to the target cell.");

        IXlsxWorksheet targetSheet = targetCell.Worksheet;

        var pivotTableName = this.Name;

        int i = 0;
        var pivotTableNames = targetSheet.PivotTables.Select(pvt => pvt.Name).ToList();
        while (!XlsxHelper.ValidateName("pivot table", pivotTableName, "", pivotTableNames, out _))
        {
            i++;
            pivotTableName = Name + i.ToInvariantString();
        }

        var newPivotTable = (XlsxPivotTable)targetSheet.PivotTables.Add(pivotTableName, targetCell, PivotCache);

        newPivotTable.RelId = null;

        static void CopyPivotField(IXlsxPivotField originalPivotField, IXlsxPivotField newPivotField)
        {
            newPivotField
                .SetSort(originalPivotField.SortType)
                .SetSubtotalCaption(originalPivotField.SubtotalCaption)
                .SetIncludeNewItemsInFilter(originalPivotField.IncludeNewItemsInFilter)
                .SetRepeatItemLabels(originalPivotField.RepeatItemLabels)
                .SetInsertBlankLines(originalPivotField.InsertBlankLines)
                .SetShowBlankItems(originalPivotField.ShowBlankItems)
                .SetInsertPageBreaks(originalPivotField.InsertPageBreaks)
                .SetCollapsed(originalPivotField.Collapsed);

            if (originalPivotField.SubtotalsAtTop.HasValue)
                newPivotField.SetSubtotalsAtTop(originalPivotField.SubtotalsAtTop.Value);

            newPivotField.AddSelectedValues(originalPivotField.SelectedValues);
        }

        foreach (IXlsxPivotField rf in ReportFilters)
            CopyPivotField(rf, newPivotTable.ReportFilters.Add(rf.SourceName, rf.CustomName));

        foreach (IXlsxPivotField cl in ColumnLabels)
            CopyPivotField(cl, newPivotTable.ColumnLabels.Add(cl.SourceName, cl.CustomName));

        foreach (IXlsxPivotField rl in RowLabels)
            CopyPivotField(rl, newPivotTable.RowLabels.Add(rl.SourceName, rl.CustomName));

        foreach (IXlsxPivotValue v in Values)
        {
            IXlsxPivotValue pivotValue = newPivotTable.Values.Add(v.SourceName, v.CustomName)
                .SetSummaryFormula(v.SummaryFormula)
                .SetCalculation(v.Calculation)
                .SetCalculationItem(v.CalculationItem)
                .SetBaseFieldName(v.BaseFieldName)
                .SetBaseItemValue(v.BaseItemValue);

            pivotValue.NumberFormat.NumberFormatId = v.NumberFormat.NumberFormatId;
            pivotValue.NumberFormat.Format = v.NumberFormat.Format;
        }

        newPivotTable.Title = Title;
        newPivotTable.Description = Description;
        newPivotTable.ColumnHeaderCaption = ColumnHeaderCaption;
        newPivotTable.RowHeaderCaption = RowHeaderCaption;
        newPivotTable.MergeAndCenterWithLabels = MergeAndCenterWithLabels;
        newPivotTable.RowLabelIndent = RowLabelIndent;
        newPivotTable.FilterAreaOrder = FilterAreaOrder;
        newPivotTable.FilterFieldsPageWrap = FilterFieldsPageWrap;
        newPivotTable.ErrorValueReplacement = ErrorValueReplacement;
        newPivotTable.ShowMissing = ShowMissing;
        newPivotTable.MissingCaption = MissingCaption;
        newPivotTable.AutofitColumns = AutofitColumns;
        newPivotTable.PreserveCellFormatting = PreserveCellFormatting;
        newPivotTable.ShowGrandTotalsColumns = ShowGrandTotalsColumns;
        newPivotTable.ShowGrandTotalsRows = ShowGrandTotalsRows;
        newPivotTable.FilteredItemsInSubtotals = FilteredItemsInSubtotals;
        newPivotTable.AllowMultipleFilters = AllowMultipleFilters;
        newPivotTable.UseCustomListsForSorting = UseCustomListsForSorting;
        newPivotTable.ShowExpandCollapseButtons = ShowExpandCollapseButtons;
        newPivotTable.ShowContextualTooltips = ShowContextualTooltips;
        newPivotTable.ShowPropertiesInTooltips = ShowPropertiesInTooltips;
        newPivotTable.DisplayCaptionsAndDropdowns = DisplayCaptionsAndDropdowns;
        newPivotTable.ClassicPivotTableLayout = ClassicPivotTableLayout;
        newPivotTable.ShowValuesRow = ShowValuesRow;
        newPivotTable.ShowEmptyItemsOnColumns = ShowEmptyItemsOnColumns;
        newPivotTable.ShowEmptyItemsOnRows = ShowEmptyItemsOnRows;
        newPivotTable.DisplayItemLabels = DisplayItemLabels;
        newPivotTable.SortFieldsAtoZ = SortFieldsAtoZ;
        newPivotTable.PrintExpandCollapsedButtons = PrintExpandCollapsedButtons;
        newPivotTable.RepeatRowLabels = RepeatRowLabels;
        newPivotTable.PrintTitles = PrintTitles;
        newPivotTable.EnableShowDetails = EnableShowDetails;
        newPivotTable.EnableCellEditing = EnableCellEditing;
        newPivotTable.ShowRowHeaders = ShowRowHeaders;
        newPivotTable.ShowColumnHeaders = ShowColumnHeaders;
        newPivotTable.ShowRowStripes = ShowRowStripes;
        newPivotTable.ShowColumnStripes = ShowColumnStripes;
        newPivotTable.Theme = Theme;
        // TODO: Copy Styleformats

        return newPivotTable;
    }

    public string Name
    {
        get { return _name; }
        set
        {
            if (_name == value) return;

            var oldname = _name ?? string.Empty;

            if (!XlsxHelper.ValidateName("pivot table", value, oldname, _worksheet.PivotTables.Select<XlsxPivotTable, string>(pvt => pvt.Name), out var message))
                throw new ArgumentException(message, nameof(value));

            _name = value;

            if (!string.IsNullOrWhiteSpace(oldname) && !string.Equals(oldname, _name, StringComparison.OrdinalIgnoreCase))
            {
                Worksheet.PivotTables.Delete(oldname);
                (Worksheet.PivotTables as XlsxPivotTables).Add(_name, this);
            }
        }
    }

    public IXlsxPivotTable SetName(string value)
    {
        Name = value; return this;
    }

    public string Title { get; set; }

    public IXlsxPivotTable SetTitle(string value)
    {
        Title = value; return this;
    }

    public string Description { get; set; }

    public IXlsxPivotTable SetDescription(string value)
    {
        Description = value; return this;
    }

    public IXlsxPivotTable SetColumnHeaderCaption(string value)
    {
        ColumnHeaderCaption = value;
        return this;
    }

    public IXlsxPivotTable SetRowHeaderCaption(string value)
    {
        RowHeaderCaption = value;
        return this;
    }

    public IXlsxPivotTable SetMergeAndCenterWithLabels()
    {
        MergeAndCenterWithLabels = true; return this;
    }

    public IXlsxPivotTable SetMergeAndCenterWithLabels(bool value)
    {
        MergeAndCenterWithLabels = value; return this;
    }

    public IXlsxPivotTable SetRowLabelIndent(int value)
    {
        RowLabelIndent = value; return this;
    }

    public IXlsxPivotTable SetFilterAreaOrder(XlsxFilterAreaOrder value)
    {
        FilterAreaOrder = value; return this;
    }

    public IXlsxPivotTable SetFilterFieldsPageWrap(int value)
    {
        FilterFieldsPageWrap = value; return this;
    }

    public IXlsxPivotTable SetErrorValueReplacement(string value)
    {
        ErrorValueReplacement = value; return this;
    }

    public string EmptyCellReplacement
    {
        get
        {
            if (ShowMissing)
                return MissingCaption;

            return string.Empty;
        }
        set
        {
            if (string.IsNullOrEmpty(value))
            {
                ShowMissing = false;
                MissingCaption = string.Empty;
            }
            else
            {
                ShowMissing = true;
                MissingCaption = value;
            }
        }
    }

    public IXlsxPivotTable SetEmptyCellReplacement(string value)
    {
        EmptyCellReplacement = value; return this;
    }

    public IXlsxPivotTable SetAutofitColumns()
    {
        AutofitColumns = true; return this;
    }

    public IXlsxPivotTable SetAutofitColumns(bool value)
    {
        AutofitColumns = value; return this;
    }

    public IXlsxPivotTable SetPreserveCellFormatting()
    {
        PreserveCellFormatting = true; return this;
    }

    public IXlsxPivotTable SetPreserveCellFormatting(bool value)
    {
        PreserveCellFormatting = value; return this;
    }

    public IXlsxPivotTable SetShowGrandTotalsRows()
    {
        ShowGrandTotalsRows = true; return this;
    }

    public IXlsxPivotTable SetShowGrandTotalsRows(bool value)
    {
        ShowGrandTotalsRows = value; return this;
    }

    public IXlsxPivotTable SetShowGrandTotalsColumns()
    {
        ShowGrandTotalsColumns = true; return this;
    }

    public IXlsxPivotTable SetShowGrandTotalsColumns(bool value)
    {
        ShowGrandTotalsColumns = value; return this;
    }

    public IXlsxPivotTable SetFilteredItemsInSubtotals()
    {
        FilteredItemsInSubtotals = true; return this;
    }

    public IXlsxPivotTable SetFilteredItemsInSubtotals(bool value)
    {
        FilteredItemsInSubtotals = value; return this;
    }

    public IXlsxPivotTable SetAllowMultipleFilters()
    {
        AllowMultipleFilters = true; return this;
    }

    public IXlsxPivotTable SetAllowMultipleFilters(bool value)
    {
        AllowMultipleFilters = value; return this;
    }

    public IXlsxPivotTable SetUseCustomListsForSorting()
    {
        UseCustomListsForSorting = true; return this;
    }

    public IXlsxPivotTable SetUseCustomListsForSorting(bool value)
    {
        UseCustomListsForSorting = value; return this;
    }

    public IXlsxPivotTable SetShowExpandCollapseButtons()
    {
        ShowExpandCollapseButtons = true; return this;
    }

    public IXlsxPivotTable SetShowExpandCollapseButtons(bool value)
    {
        ShowExpandCollapseButtons = value; return this;
    }

    public IXlsxPivotTable SetShowContextualTooltips()
    {
        ShowContextualTooltips = true; return this;
    }

    public IXlsxPivotTable SetShowContextualTooltips(bool value)
    {
        ShowContextualTooltips = value; return this;
    }

    public IXlsxPivotTable SetShowPropertiesInTooltips()
    {
        ShowPropertiesInTooltips = true; return this;
    }

    public IXlsxPivotTable SetShowPropertiesInTooltips(bool value)
    {
        ShowPropertiesInTooltips = value; return this;
    }

    public IXlsxPivotTable SetDisplayCaptionsAndDropdowns()
    {
        DisplayCaptionsAndDropdowns = true; return this;
    }

    public IXlsxPivotTable SetDisplayCaptionsAndDropdowns(bool value)
    {
        DisplayCaptionsAndDropdowns = value; return this;
    }

    public IXlsxPivotTable SetClassicPivotTableLayout()
    {
        ClassicPivotTableLayout = true; return this;
    }

    public IXlsxPivotTable SetClassicPivotTableLayout(bool value)
    {
        ClassicPivotTableLayout = value; return this;
    }

    public bool ShowValuesRow { get; set; }

    public IXlsxPivotTable SetShowValuesRow()
    {
        ShowValuesRow = true; return this;
    }

    public IXlsxPivotTable SetShowValuesRow(bool value)
    {
        ShowValuesRow = value; return this;
    }

    public IXlsxPivotTable SetShowEmptyItemsOnRows()
    {
        ShowEmptyItemsOnRows = true; return this;
    }

    public IXlsxPivotTable SetShowEmptyItemsOnRows(bool value)
    {
        ShowEmptyItemsOnRows = value; return this;
    }

    public IXlsxPivotTable SetShowEmptyItemsOnColumns()
    {
        ShowEmptyItemsOnColumns = true; return this;
    }

    public IXlsxPivotTable SetShowEmptyItemsOnColumns(bool value)
    {
        ShowEmptyItemsOnColumns = value; return this;
    }

    public IXlsxPivotTable SetDisplayItemLabels()
    {
        DisplayItemLabels = true; return this;
    }

    public IXlsxPivotTable SetDisplayItemLabels(bool value)
    {
        DisplayItemLabels = value; return this;
    }

    public IXlsxPivotTable SetSortFieldsAtoZ()
    {
        SortFieldsAtoZ = true; return this;
    }

    public IXlsxPivotTable SetSortFieldsAtoZ(bool value)
    {
        SortFieldsAtoZ = value; return this;
    }

    public IXlsxPivotTable SetPrintExpandCollapsedButtons()
    {
        PrintExpandCollapsedButtons = true; return this;
    }

    public IXlsxPivotTable SetPrintExpandCollapsedButtons(bool value)
    {
        PrintExpandCollapsedButtons = value; return this;
    }

    public IXlsxPivotTable SetRepeatRowLabels()
    {
        RepeatRowLabels = true; return this;
    }

    public IXlsxPivotTable SetRepeatRowLabels(bool value)
    {
        RepeatRowLabels = value; return this;
    }

    public IXlsxPivotTable SetPrintTitles()
    {
        PrintTitles = true; return this;
    }

    public IXlsxPivotTable SetPrintTitles(bool value)
    {
        PrintTitles = value; return this;
    }

    public IXlsxPivotTable SetEnableShowDetails()
    {
        EnableShowDetails = true; return this;
    }

    public IXlsxPivotTable SetEnableShowDetails(bool value)
    {
        EnableShowDetails = value; return this;
    }


    public bool EnableCellEditing { get; set; }

    public IXlsxPivotTable SetEnableCellEditing()
    {
        EnableCellEditing = true; return this;
    }

    public IXlsxPivotTable SetEnableCellEditing(bool value)
    {
        EnableCellEditing = value; return this;
    }

    public bool ShowRowHeaders { get; set; }

    public IXlsxPivotTable SetShowRowHeaders()
    {
        ShowRowHeaders = true; return this;
    }

    public IXlsxPivotTable SetShowRowHeaders(bool value)
    {
        ShowRowHeaders = value; return this;
    }

    public bool ShowColumnHeaders { get; set; }

    public IXlsxPivotTable SetShowColumnHeaders()
    {
        ShowColumnHeaders = true; return this;
    }

    public IXlsxPivotTable SetShowColumnHeaders(bool value)
    {
        ShowColumnHeaders = value; return this;
    }

    public bool ShowRowStripes { get; set; }

    public IXlsxPivotTable SetShowRowStripes()
    {
        ShowRowStripes = true; return this;
    }

    public IXlsxPivotTable SetShowRowStripes(bool value)
    {
        ShowRowStripes = value; return this;
    }

    public bool ShowColumnStripes { get; set; }

    public IXlsxPivotTable SetShowColumnStripes()
    {
        ShowColumnStripes = true; return this;
    }

    public IXlsxPivotTable SetShowColumnStripes(bool value)
    {
        ShowColumnStripes = value; return this;
    }

    /// <summary>
    /// Part of the pivot table style.
    /// </summary>
    internal bool ShowLastColumn { get; set; } = false;

    public XlsxPivotSubtotals Subtotals { get; set; }

    public IXlsxPivotTable SetSubtotals(XlsxPivotSubtotals value)
    {
        Subtotals = value; return this;
    }

    public XlsxPivotLayout Layout
    {
        set
        {
            switch (value)
            {
                case XlsxPivotLayout.Compact:
                    _compact = _compactData = true;
                    _outline = _outlineData = false;
                    break;
                case XlsxPivotLayout.Outline:
                    _compact = _compactData = false;
                    _outline = _outlineData = true;
                    break;
                case XlsxPivotLayout.Tabular:
                    _compact = _compactData = false;
                    _outline = _outlineData = false;
                    break;
                default:
                    throw new UnreachableException();
            }

            // It is necessary to set layout for each pivot field, even ones that are not displayed on an axis. Without it, the tabular layout
            // doesn't display headers for axis fields and only display one "Column labels" button instead.
            PivotFields.ForEach(f => f.SetLayout(value));
        }
    }

    public IXlsxPivotTable SetLayout(XlsxPivotLayout value)
    {
        Layout = value; return this;
    }

    public bool InsertBlankLines
    {
        set { ImplementedFields.ForEach(f => f.SetInsertBlankLines(value)); }
    }

    public IXlsxPivotTable SetInsertBlankLines()
    {
        InsertBlankLines = true; return this;
    }

    public IXlsxPivotTable SetInsertBlankLines(bool value)
    {
        InsertBlankLines = value; return this;
    }

    internal string RelId { get; set; }
    internal string CacheDefinitionRelId { get; set; }

    private void SetExcelDefaults()
    {
        ShowMissing = true;
        MissingCaption = string.Empty;
        ShowColumnHeaders = true;
        ShowRowHeaders = true;

        // source http://www.datypic.com/sc/ooxml/e-ssml_pivotTableDefinition.html
        DisplayItemLabels = true;	//	Show Item Names
        ShowExpandCollapseButtons = true;	//	Show Expand Collapse
        PrintExpandCollapsedButtons = false;	//	Print Drill Indicators
        ShowPropertiesInTooltips = true;	//	Show Member Property ToolTips
        ShowContextualTooltips = true;	//	Show ToolTips on Data
        EnableShowDetails = true;	//	Enable Drill Down
        PreserveCellFormatting = true;	//	Preserve Formatting
        AutofitColumns = false;	//	Auto Formatting
        FilterAreaOrder = XlsxFilterAreaOrder.DownThenOver;	//	Page Over Then Down
        FilteredItemsInSubtotals = false;	//	Subtotal Hidden Items
        ShowGrandTotalsRows = true;	//	Row Grand Totals
        ShowGrandTotalsColumns = true;	//	Grand Totals On Columns
        PrintTitles = false;	//	Field Print Titles
        RepeatRowLabels = false;	//	Item Print Titles
        MergeAndCenterWithLabels = false;	//	Merge Titles
        RowLabelIndent = 1;	//	Indentation for Compact Axis
        ShowEmptyItemsOnRows = false;	//	Show Empty Row
        ShowEmptyItemsOnColumns = false;	//	Show Empty Column
        DisplayCaptionsAndDropdowns = true;	//	Show Field Headers
        ClassicPivotTableLayout = false;	//	Enable Drop Zones
        AllowMultipleFilters = true;	//	Multiple Field Filters
        SortFieldsAtoZ = false;	//	Default Sort Order
        UseCustomListsForSorting = true; //	Custom List AutoSort
    }

    public IXlsxWorksheet Worksheet => _worksheet;

    public IXlsxPivotTableStyleFormats StyleFormats => new XlsxPivotTableStyleFormats(this);

    public IEnumerable<IXlsxPivotStyleFormat> AllStyleFormats
    {
        get
        {
            foreach (IXlsxPivotStyleFormat styleFormat in this.StyleFormats.RowGrandTotalFormats)
                yield return styleFormat;

            foreach (IXlsxPivotStyleFormat styleFormat in this.StyleFormats.ColumnGrandTotalFormats)
                yield return styleFormat;

            // TODO: Skipped for now, until I implement stubs
            //foreach (var pivotField in ImplementedFields)
            //{
            //    yield return pivotField.StyleFormats.Subtotal;
            //    yield return pivotField.StyleFormats.Header;
            //    yield return pivotField.StyleFormats.Label;
            //    yield return pivotField.StyleFormats.DataValuesFormat;
            //}
        }
    }
#nullable enable
    internal void AddField(XlsxPivotTableField field)
    {
        _fields.Add(field);
    }

    internal void AddFormat(XlsxPivotFormat pivotFormat)
    {
        _formats.Add(pivotFormat);
    }

    internal void AddConditionalFormat(XlsxPivotConditionalFormat conditionalFormat)
    {
        _conditionalFormats.Add(conditionalFormat);
    }

    #region location

    /// <summary>
    /// Area of a pivot table. Area doesn't include page fields, they are above the area with
    /// one empty row between area and filters.
    /// </summary>
    internal Area Area { get; set; } = new(1, 1, 1, 1);

    /// <summary>
    /// First row of pivot table header, relative to the <see cref="Area"/>.
    /// </summary>
    internal uint FirstHeaderRow { get; set; }

    /// <summary>
    /// First row of pivot table data area, relative to the <see cref="Area"/>.
    /// </summary>
    internal uint FirstDataRow { get; set; }

    /// <summary>
    /// First column of pivot table data area, relative to the <see cref="Area"/>.
    /// </summary>
    internal uint FirstDataCol { get; set; }

    #endregion

    #region Attributes of PivotTableDefinition in same order as XSD

    /// <summary>
    /// Determines the whether 'data' field is on <see cref="RowAxis"/> (<c>true</c>) or
    /// <see cref="ColumnAxis"/>(<c>false</c>).
    /// </summary>
    internal bool DataOnRows { get; set; } = false;

    /// <summary>
    /// Determines the default 'data' field position, when it is automatically added to row/column fields.
    /// 0 = first (e.g. before all column/row fields), 1 = second (i.e. after first row/column field) and so on.
    /// &gt; number of fields or <c>null</c> indicates the last position.
    /// </summary>
    internal int? DataPosition { get; set; }

    /// <summary>
    /// <para>
    /// An identification of legacy table auto-format to apply to the pivot table. The
    /// <c>Apply*Formats</c> properties specifies which parts of auto-format to apply. If
    /// <c>null</c> or <see cref="AutofitColumns"/> is not <c>true</c>, legacy auto-format is
    /// not applied.
    /// </para>
    /// <para>
    /// The value must be less than 21 or greater than 4096 and less than or equal to 4117. See
    /// ISO-29500 Annex G.3 for how auto formats look like.
    /// </para>
    /// </summary>
    internal uint? AutoFormatId { get; init; }

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format number format properties.
    /// </summary>
    internal bool ApplyNumberFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format border properties.
    /// </summary>
    internal bool ApplyBorderFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format font properties.
    /// </summary>
    internal bool ApplyFontFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format pattern properties.
    /// </summary>
    internal bool ApplyPatternFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format alignment properties.
    /// </summary>
    internal bool ApplyAlignmentFormats { get; init; } = false;

    /// <summary>
    /// If auto-format should be applied (<see cref="AutofitColumns"/> and <see cref="AutoFormatId"/>
    /// are set), apply legacy auto-format width/height properties.
    /// </summary>
    internal bool ApplyWidthHeightFormats { get; init; } = false;

    /// <summary>
    /// Initial text of 'data' field. This is doesn't do anything, Excel always displays
    /// dynamically a text 'Values', translated to current culture.
    /// </summary>
    internal string DataCaption { get; set; } = "Values";

    internal string? GrandTotalCaption { get; init; }

    /// <summary>
    /// Text to display when in cells that contain error.
    /// </summary>
    public string? ErrorValueReplacement { get; set; }

    /// <summary>
    /// Flag indicating if <see cref="ErrorValueReplacement"/> should be shown when cell contain an error.
    /// </summary>
    internal bool ShowError { get; init; } = false;

    /// <summary>
    /// Test to display for missing items, when <see cref="ShowMissing"/> is <c>true</c>.
    /// </summary>
    internal string MissingCaption { get; set; }

    /// <summary>
    /// Flag indicating if <see cref="MissingCaption"/> should be shown when cell has no value.
    /// </summary>
    /// <remarks>Doesn't seem to work in Excel.</remarks>
    internal bool ShowMissing { get; set; } = true;

    /// <summary>
    /// Name of style to apply to <see cref="XlsxPivotPageField"/> items headers in <see cref="XlsxPivotAxis.AxisPage"/>.
    /// </summary>
    internal string? PageStyle { get; init; }

    /// <remarks>Doesn't seem to work in Excel.</remarks>
    internal string? PivotTableStyleName { get; init; }

    /// <summary>
    /// Name of a style to apply to the cells left blank when a pivot table shrinks during a refresh operation.
    /// </summary>
    internal string? VacatedStyle { get; init; }

    internal string? Tag { get; init; }

    /// <summary>
    /// Version of the application that last updated the pivot table. Application-dependent.
    /// </summary>
    internal byte UpdatedVersion { get; init; }

    /// <summary>
    /// Minimum version of the application required to update the pivot table. Application-dependent.
    /// </summary>
    internal byte MinRefreshableVersion { get; init; }

    /// <remarks>OLAP related.</remarks>
    internal bool AsteriskTotals { get; init; } = false;

    /// <summary>
    /// <para>
    /// Should field items be displayed on the axis despite pivot table not having any value
    /// field? <c>true</c> will display items even without data field, <c>false</c> won't.
    /// </para>
    /// <para>
    /// Example: There is an empty pivot table with no value fields. Add field 'Name'
    /// to row fields. Should names be displayed on row despite not having any value field?
    /// </para>
    /// </summary>
    /// <remarks>Also called ShowItems</remarks>
    public bool DisplayItemLabels { get; set; } = true;

    /// <summary>
    /// Flag indicating if user is allowed to edit cells in data area.
    /// </summary>
    internal bool EditData { get; init; } = false;

    /// <summary>
    /// Flag indicating if UI to modify the fields of pivot table is disabled. In Excel, the
    /// whole field area is hidden.
    /// </summary>
    internal bool DisableFieldList { get; init; } = false;

    /// <remarks>OLAP only.</remarks>
    internal bool ShowCalculatedMembers { get; init; } = true;

    /// <remarks>OLAP only.</remarks>
    internal bool VisualTotals { get; init; } = true;

    /// <summary>
    /// A flag indicating whether a page field that has selected multiple items (but not
    /// necessarily all) display "(multiple items)" instead of "All"? If value is <c>false</c>.
    /// page fields will display "All" regardless of whether only item subset is selected or
    /// all items are selected.
    /// </summary>
    internal bool ShowMultipleLabel { get; init; } = true;

    /// <summary>
    /// Doesn't seem to do anything. Should hide drop down filters.
    /// </summary>
    internal bool ShowDataDropDown { get; init; } = true;

    /// <summary>
    /// A flag indicating whether UI should display collapse/expand (drill) buttons in pivot
    /// table axes.
    /// </summary>
    /// <remarks>Also called ShowDrill.</remarks>
    public bool ShowExpandCollapseButtons { get; set; } = true;

    /// <summary>
    /// A flag indicating whether collapse/expand (drill) buttons in pivot table axes should
    /// be printed.
    /// </summary>
    /// <remarks>Also called PrintDrill.</remarks>
    public bool PrintExpandCollapsedButtons { get; set; } = false;

    /// <remarks>OLAP only. Also called ShowMemberPropertyTips.</remarks>
    public bool ShowPropertiesInTooltips { get; set; }

    /// <summary>
    /// A flag indicating whether UI should display a tooltip on data items of pivot table. The
    /// tooltip contain info about value field name, row/col items used to aggregate the value
    /// ect. Note that this tooltip generally hides cell notes, because mouseover displays data
    /// tool tip, rather than the note.
    /// </summary>
    /// <remarks>Also called ShowDataTips.</remarks>
    public bool ShowContextualTooltips { get; set; }

    /// <summary>
    /// A flag indicating whether UI should provide a mechanism to edit the pivot table. If the
    /// value is <c>false</c>, Excel provides ability to refresh data through context menu, but
    /// ribbon or other options to manipulate field or pivot table settings are not present.
    /// </summary>
    /// <remarks>Also called enableWizard.</remarks>
    internal bool EnableEditingMechanism { get; set; } = true;

    /// <remarks>Likely OLAP only. Do not confuse with collapse/expand buttons.</remarks>
    public bool EnableShowDetails { get; set; } = true;

    /// <summary>
    /// A flag indicating whether the user is prevented from displaying PivotField properties.
    /// Not very consistent in Excel, e.g. can't display field properties through context menu
    /// of a pivot table, but can display properties menu through context menu in editing wizard.
    /// </summary>
    internal bool EnableFieldProperties { get; init; } = true;

    /// <summary>
    /// A flag that indicates whether the formatting applied by the user to the pivot table
    /// cells is preserved on refresh.
    /// </summary>
    /// <remarks>Once again, ISO-29500 is buggy and says the opposite. Also called <em>
    /// PreserveFormatting</em></remarks>
    public bool PreserveCellFormatting { get; set; } = true;

    /// <summary>
    /// A flag that indicates whether legacy auto formatting has been applied to the PivotTable
    /// view.
    /// </summary>
    /// <remarks>Also called UseAutoFormatting.</remarks>
    public bool AutofitColumns { get; set; } = false;

    /// <inheritdoc />
    /// <remarks>Also called PageWrap.</remarks>
    public int FilterFieldsPageWrap
    {
        get => _filterFieldsPageWrap;
        set
        {
            ArgumentOutOfRangeException.ThrowIfNegative(value);

            _filterFieldsPageWrap = value;
        }
    }

    /// <inheritdoc />
    /// <remarks>Also called <em>PageOverThenDown</em>.</remarks>
    public XlsxFilterAreaOrder FilterAreaOrder { get; set; } = XlsxFilterAreaOrder.DownThenOver;

    /// <summary>
    /// A flag that indicates whether hidden pivot items should be included in subtotal
    /// calculated values. If <c>true</c>, data for hidden items are included in subtotals
    /// calculated values. If <c>false</c>, hidden values are not included in subtotal
    /// calculations.
    /// </summary>
    /// <remarks>Also called <em>SubtotalHiddenItems</em>. OLAP only. Option in Excel is grayed
    ///     out and does nothing. The option is un-grayed out when pivot cache is part of data
    ///     model.</remarks>
    public bool FilteredItemsInSubtotals { get; set; } = false;

    /// <summary>
    /// A flag indicating whether grand totals should be displayed for the PivotTable rows.
    /// </summary>
    /// <remarks>Also called <em>RowGrandTotals</em>.</remarks>
    public bool ShowGrandTotalsRows { get; set; } = true;

    /// <summary>
    /// A flag indicating whether grand totals should be displayed for the PivotTable columns.
    /// </summary>
    /// <remarks>Also called <em>ColumnGrandTotals</em>.</remarks>
    public bool ShowGrandTotalsColumns { get; set; } = true;

    /// <summary>
    /// A flag indicating whether when a field name should be printed on all pages.
    /// </summary>
    /// <remarks>Also called <em>FieldPrintTitles</em>.</remarks>
    public bool PrintTitles { get; set; } = false;

    /// <summary>
    /// A flag indicating whether whether PivotItem names should be repeated at the top of each
    /// printed page (e.g. if axis item spans multiple pages, it will be repeated an all pages).
    /// </summary>
    /// <remarks>Also called <em>ItemPrintTitles</em>.</remarks>
    public bool RepeatRowLabels { get; set; } = false;

    /// <summary>
    /// A flag indicating whether row or column titles that span multiple cells should be
    /// merged into a single cell. Useful only in in tabular layout, titles in other layouts
    /// don't span across multiple cells.
    /// </summary>
    /// <remarks>Also called <em>MergeItem</em>.</remarks>
    public bool MergeAndCenterWithLabels { get; set; } = false;

    /// <summary>
    /// A flag indicating whether UI for the pivot table should display large text in field
    /// drop zones when there are no fields in the data region (e.g. <em>Drop Value Fields
    /// Here</em>). Only works in legacy layout mode (i.e. <see cref="ClassicPivotTableLayout"/>
    /// is <c>true</c>).
    /// </summary>
    internal bool ShowDropZones { get; init; } = true;

    /// <summary>
    /// Specifies the version of the application that created the pivot cache. Application-dependent.
    /// </summary>
    /// <remarks>Also called <em>CreatedVersion</em>.</remarks>
    internal byte PivotCacheCreatedVersion { get; init; } = 0;

    /// <summary>
    /// A row indentation increment for row axis when pivot table is in compact layout. Units
    /// are characters.
    /// </summary>
    /// <remarks>Also called <em>Indent</em>.</remarks>
    public int RowLabelIndent { get; set; } = 1;

    /// <summary>
    /// A flag indicating whether to include empty rows in the pivot table (i.e. row axis items
    /// are blank and data items are blank).
    /// </summary>
    /// <remarks>Also called <em>ShowEmptyRow</em>.</remarks>
    public bool ShowEmptyItemsOnRows { get; set; } = false;

    /// <summary>
    /// A flag indicating whether to include empty columns in the table (i.e. column axis items
    /// are blank and data items are blank).
    /// </summary>
    /// <remarks>Also called <em>ShowEmptyColumn</em>.</remarks>
    public bool ShowEmptyItemsOnColumns { get; set; }

    /// <summary>
    /// A flag indicating whether to show field names on axis. The axis items are still
    /// displayed, only field names are not. The dropdowns next to the axis field names
    /// are also displayed/hidden based on the flag.
    /// </summary>
    /// <remarks>Also called <em>ShowHeaders</em>.</remarks>
    public bool DisplayCaptionsAndDropdowns { get; set; } = true;

    /// <summary>
    /// A flag indicating whether new fields should have their
    /// <see cref="XlsxPivotTableField.Compact"/> flag set to <c>true</c>. By new, it means field
    /// added to page, axes or data fields, not a new field from cache.
    /// </summary>
    internal bool Compact
    {
        get => _compact;
        init => _compact = value;
    }

    /// <summary>
    /// A flag indicating whether new fields should have their
    /// <see cref="XlsxPivotTableField.Outline"/> flag set to <c>true</c>. By new, it means field
    /// added to page, axes or data fields, not a new field from cache.
    /// </summary>
    internal bool Outline
    {
        get => _outline;
        init => _outline = value;
    }

    /// <summary>
    /// <para>
    /// A flag that indicates whether 'data'/-2 fields in the PivotTable should be displayed in
    /// outline next column of the sheet. This is basically an equivalent of
    /// <see cref="XlsxPivotTableField.Outline"/> property for the 'data' fields, because 'data'
    /// field is implicit.
    /// </para>
    /// <para>
    /// When <c>true</c>, the labels from the next field (as ordered by
    /// <see cref="XlsxPivotTableAxis.Fields"/> for row or column) are displayed in the next
    /// column. Has no effect if 'data' field is last field.
    /// </para>
    /// </summary>
    /// <remarks>Doesn't seem to do much in column axis, only in row axis. Also, Excel
    ///     sometimes seems to favor <see cref="Outline"/> flag instead (likely some less used
    ///     paths in the Excel code).</remarks>
    internal bool OutlineData
    {
        get => _outlineData;
        init => _outlineData = value;
    }

    /// <summary>
    /// <para>
    /// A flag that indicates whether 'data'/-2 fields in the PivotTable should be displayed in
    /// compact mode (=same column of the sheet). This is basically an equivalent of
    /// <see cref="XlsxPivotTableField.Compact"/> property for the 'data' fields, because 'data'
    /// field is implicit.
    /// </para>
    /// <para>
    /// When <c>true</c>, the labels from the next field (as ordered by
    /// <see cref="XlsxPivotTableAxis.Fields"/> for row or column) are displayed in the same
    /// column (one row below). Has no effect if 'data' field is last field.
    /// </para>
    /// </summary>
    /// <remarks>Doesn't seem to do much in column axis, only in row axis. Also, Excel
    ///     sometimes seems to favor <see cref="Compact"/> flag instead (likely some less used
    ///     paths in the Excel code).</remarks>
    internal bool CompactData
    {
        get => _compactData;
        init => _compactData = value;
    }

    /// <summary>
    /// A flag that indicates whether data fields in the pivot table are published and
    /// available for viewing in a server rendering environment.
    /// </summary>
    /// <remarks>No idea what this does. Likely flag for other components that display table
    ///     on a web page.</remarks>
    internal bool Published { get; init; } = false;

    /// <summary>
    /// A flag that indicates whether to apply the classic layout. Classic layout displays the
    /// grid zones in UI where user can drop fields (unless disabled through
    /// <see cref="ShowDropZones"/>).
    /// </summary>
    /// <remarks>Also called <em>GridDropZones</em>.</remarks>
    public bool ClassicPivotTableLayout { get; set; } = false;

    /// <summary>
    /// Likely a flag whether immersive reader should be turned off. Not sure if immersive
    /// reader was ever used outside Word, though Excel for Web added some support in 2023.
    /// </summary>
    internal bool StopImmersiveUi { get; init; } = true;

    /// <summary>
    /// <para>
    /// A flag indicating whether field can have at most most one filter type used. This flag
    /// doesn't allow multiple filters of same type, only multiple different filter types.
    /// </para>
    /// <para>
    /// If false, field can have at most one filter, if user tries to set multiple, previous
    /// one is cleared.
    /// </para>
    /// </summary>
    /// <remarks>Also called <em>multipleFieldFilters</em>.</remarks>
    public bool AllowMultipleFilters { get; set; } = true;

    /// <summary>
    /// Specifies the next pivot chart formatting identifier to use on the pivot table. First
    /// actually used identifier should be 1. The format is used in <c>/chartSpace/pivotSource/
    /// fmtId/@val</c>.
    /// </summary>
    internal uint ChartFormat { get; init; } = 0;

    /// <summary>
    /// The text that will be displayed in row header in compact mode. It is next to drop down
    /// (if enabled) of a label/values filter for fields (if
    /// <see cref="DisplayCaptionsAndDropdowns"/> is set to <c>true</c>). Use localized text
    /// <em>Row labels</em> if property is not specified.
    /// </summary>
    public string? RowHeaderCaption { get; set; } = null;

    /// <summary>
    /// The text that will be displayed in column header in compact mode. It is next to drop down
    /// (if enabled) of a label/values filter for fields (if
    /// <see cref="DisplayCaptionsAndDropdowns"/> is set to <c>true</c>). Use localized text
    /// <em>Column labels</em> if property is not specified.
    /// </summary>
    public string? ColumnHeaderCaption { get; set; } = null;

    /// <summary>
    /// A flag that controls how are fields sorted in the field list UI. <c>true</c> will
    /// display fields sorted alphabetically, <c>false</c> will display fields in the order
    /// fields appear in <see cref="XlsxPivotCache"/>. OLAP data sources always use alphabetical
    /// sorting.
    /// </summary>
    /// <remarks>Also called <em>fieldListSortAscending</em>.</remarks>
    public bool SortFieldsAtoZ { get; set; } = false;

    /// <summary>
    /// A flag indicating whether MDX sub-queries are supported by OLAP data provider of this
    /// pivot table.
    /// </summary>
    internal bool MdxSubQueries { get; init; } = false;

    /// <summary>
    /// A flag that indicates whether custom lists are used for sorting items of fields, both
    /// initially when the PivotField is initialized and the PivotItems are ordered by their
    /// captions, and later when the user applies a sort.
    /// </summary>
    /// <remarks>Also called <em>customSortList</em>.</remarks>
    public bool UseCustomListsForSorting { get; set; }

    #endregion

    /// <summary>
    /// Add field to a specific axis (page/row/col). Only modified <see cref="PivotFields"/>, doesn't modify
    /// additional info in <see cref="RowAxis"/>, <see cref="ColumnAxis"/> or <see cref="Filters"/>.
    /// </summary>
    internal FieldIndex AddFieldToAxis(string sourceName, string customName, XlsxPivotAxis axis)
    {
        // Only slices axes can be added through this method.
        Debug.Assert(axis is XlsxPivotAxis.AxisCol or XlsxPivotAxis.AxisRow or XlsxPivotAxis.AxisPage);
        if (sourceName == XlsxConstants.PivotTable.ValuesSentinalLabel)
        {
            if (axis != XlsxPivotAxis.AxisRow && axis != XlsxPivotAxis.AxisCol)
                throw new ArgumentException("Data field can be used only on row or column axis.", nameof(sourceName));

            if (RowAxis.ContainsDataField || ColumnAxis.ContainsDataField)
                throw new ArgumentException("Data field is already used.", nameof(sourceName));

            var isRowAxis = axis == XlsxPivotAxis.AxisRow;

            DataOnRows = isRowAxis;
            DataPosition = isRowAxis ? RowAxis.Fields.Count : ColumnAxis.Fields.Count;
            DataCaption = "Values"; // Custom captions don't do anything.
            return FieldIndex.DataField;
        }

        if (!_cache.TryGetFieldIndex(sourceName, out var fieldIndex))
            throw new InvalidOperationException($"Field '{sourceName}' not found in pivot cache.");

        // Check actual fields.
        var customNameUsed = _fields.Any(f => XlsxHelper.NameComparer.Equals(f.Name, customName));
        if (customNameUsed)
            throw new InvalidOperationException($"Custom name '{customName}' is already used.");

        XlsxPivotTableField field = _fields[fieldIndex];
        field.Name = customName;
        field.Axis = axis;

        // If it is an axis, all possible values to field items, because they should be referenced in items.
        // Page field must have default item, otherwise Excel asks for repair.
        XlsxPivotCacheSharedItems sharedItems = _cache.GetFieldSharedItems(fieldIndex);
        for (var i = 0; i < sharedItems.Count; ++i)
            field.AddItem(new XlsxPivotFieldItem(field, i));

        // Subtotal items must be synchronized with subtotals. If field has a an item for
        // subtotal function, but doesn't declare subtotals function, Excel will try to
        // repair workbook. Subtotal items can be in any order.
        foreach (XlsxSubtotalFunction subtotalFunction in field.Subtotals)
        {
            XlsxPivotItemType itemType = subtotalFunction switch
            {
                XlsxSubtotalFunction.Automatic => XlsxPivotItemType.Default,
                XlsxSubtotalFunction.Sum => XlsxPivotItemType.Sum,
                XlsxSubtotalFunction.Count => XlsxPivotItemType.CountA,
                XlsxSubtotalFunction.Average => XlsxPivotItemType.Avg,
                XlsxSubtotalFunction.Minimum => XlsxPivotItemType.Min,
                XlsxSubtotalFunction.Maximum => XlsxPivotItemType.Max,
                XlsxSubtotalFunction.Product => XlsxPivotItemType.Product,
                XlsxSubtotalFunction.CountNumbers => XlsxPivotItemType.Count,
                XlsxSubtotalFunction.StandardDeviation => XlsxPivotItemType.StdDev,
                XlsxSubtotalFunction.PopulationStandardDeviation => XlsxPivotItemType.StdDevP,
                XlsxSubtotalFunction.Variance => XlsxPivotItemType.Var,
                XlsxSubtotalFunction.PopulationVariance => XlsxPivotItemType.VarP,
                _ => throw new UnreachableException(),
            };
            field.AddItem(new XlsxPivotFieldItem(field, null) { ItemType = itemType });
        }

        return fieldIndex;
    }

    internal void RemoveFieldFromAxis(FieldIndex index)
    {
        if (index.IsDataField)
        {
            DataOnRows = false;
            DataPosition = null;
            DataCaption = "Values";
        }
        else
        {
            XlsxPivotTableField field = _fields[index];
            field.Name = null;
            field.Axis = null;
            field.DataField = false;
            field.MultipleItemSelectionAllowed = false;
        }
    }

    internal bool TryGetSourceNameFieldIndex(string sourceName, out FieldIndex index)
    {
        if (XlsxHelper.NameComparer.Equals(sourceName, XlsxConstants.PivotTable.ValuesSentinalLabel))
        {
            index = FieldIndex.DataField;
            return true;
        }

        if (PivotCache.TryGetFieldIndex(sourceName, out var fldIndex))
        {
            index = fldIndex;
            return true;
        }

        index = default;
        return false;
    }

    internal bool TryGetCustomNameFieldIndex(string customName, out FieldIndex index)
    {
        StringComparer comparer = XlsxHelper.NameComparer;
        if (comparer.Equals(customName, XlsxConstants.PivotTable.ValuesSentinalLabel))
        {
            index = FieldIndex.DataField;
            return true;
        }

        IReadOnlyList<XlsxPivotTableField> allFields = PivotFields;
        for (var i = 0; i < allFields.Count; ++i)
        {
            if (comparer.Equals(customName, allFields[i].Name))
            {
                index = i;
                return true;
            }
        }

        index = default;
        return false;
    }

    /// <summary>
    /// Refresh cache fields after cache has changed.
    /// </summary>
    internal void UpdateCacheFields(IReadOnlyList<string> oldFieldNames)
    {
        // Should be better, but at least refresh fields. A lot of attributes are not
        // kept/initialized from the table. We can't just reuse original objects, because
        // all indices are wrong. Make a copy and then re-set the original properties that
        // are saved before GC takes them.
        var newNames = new HashSet<string>(PivotCache.FieldNames, XlsxHelper.NameComparer);

        // Source and custom name might not be valid at this point, so keep them.
        var keptDataFields = new List<(string SourceName, string? CustomName, XlsxPivotDataField Field)>();
        foreach (XlsxPivotDataField dataField in DataFields)
        {
            var oldSourceName = oldFieldNames[dataField.Field];
            if (newNames.Contains(oldSourceName))
            {
                keptDataFields.Add((oldSourceName, dataField.DataFieldName, dataField));
            }
        }

        var includeValuesField = keptDataFields.Count > 1;
        List<string> keptFilterSourceNames = GetKeptNames([.. Filters.Fields.Select(x => (FieldIndex)x.Field)], oldFieldNames, newNames, includeValuesField);
        List<string> keptRowSourceNames = GetKeptNames(RowAxis.Fields, oldFieldNames, newNames, includeValuesField);
        List<string> keptColumnSourceNames = GetKeptNames(ColumnAxis.Fields, oldFieldNames, newNames, includeValuesField);

        Filters.Clear();
        RowAxis.Clear();
        ColumnAxis.Clear();
        DataFields.Clear();

        _fields.Clear();
        foreach (var fieldName in PivotCache.FieldNames)
        {
            var field = new XlsxPivotTableField(this)
            {
                Compact = Compact,
                Outline = Outline,
            };
            _fields.Add(field);
        }

        foreach (var filterName in keptFilterSourceNames)
            Filters.Add(filterName, filterName);

        foreach (var rowName in keptRowSourceNames)
            RowAxis.AddField(rowName, rowName);

        foreach (var columnName in keptColumnSourceNames)
            ColumnAxis.AddField(columnName, columnName);

        foreach ((string? SourceName, string? CustomName, XlsxPivotDataField? Field) in keptDataFields)
        {
            XlsxPivotDataField dataField = DataFields.AddField(SourceName, CustomName);
            dataField.Subtotal = Field.Subtotal;
        }

        static List<string> GetKeptNames(
            IReadOnlyList<FieldIndex> fieldIndexes,
            IReadOnlyList<string> oldNames,
            HashSet<string> newNames,
            bool includeDataField)
        {
            var result = new List<string>();
            foreach (FieldIndex fieldIndex in fieldIndexes)
            {
                if (fieldIndex.IsDataField && includeDataField)
                {
                    result.Add(XlsxConstants.PivotTable.ValuesSentinalLabel);
                    continue;
                }

                var oldName = oldNames[fieldIndex];
                if (newNames.Contains(oldName))
                    result.Add(oldName);
            }

            return result;
        }
    }

    /// <summary>
    /// Is field used by any axis (row, column, filter), but not data.
    /// </summary>
    internal bool IsFieldUsedOnAxis(FieldIndex fieldIndex)
    {
        if (fieldIndex.IsDataField)
            return DataPosition is not null;

        return RowAxis.Fields.Contains(fieldIndex) ||
               ColumnAxis.Fields.Contains(fieldIndex) ||
               Filters.Contains(fieldIndex);
    }

    internal int GetFieldIndex(XlsxPivotTableField field)
    {
        var index = _fields.IndexOf(field);
        if (index < 0)
            throw new ArgumentException($"Unable to find field '{field.Name}'.");
        return index;
    }
}
