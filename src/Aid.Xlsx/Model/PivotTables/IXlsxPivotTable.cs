#nullable disable

using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.PivotTables.PivotStyleFormats;
using Aid.Xlsx.Model.PivotTables.PivotValues;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PivotTables;

public interface IXlsxPivotTable
{
    XlsxPivotTableTheme Theme { get; set; }

    IXlsxPivotFields ReportFilters { get; }

    /// <summary>
    /// Labels displayed in columns (i.e. horizontal axis) of the pivot table.
    /// </summary>
    IXlsxPivotFields ColumnLabels { get; }

    /// <summary>
    /// Labels displayed in rows (i.e. vertical axis) of the pivot table.
    /// </summary>
    IXlsxPivotFields RowLabels { get; }
    IXlsxPivotValues Values { get; }

    string Name { get; set; }
    string Title { get; set; }
    string Description { get; set; }

    string ColumnHeaderCaption { get; set; }
    string RowHeaderCaption { get; set; }

    /// <summary>
    /// Top left corner cell of a pivot table. If the pivot table contains filters fields, the target cell is top
    /// left cell of the first filter field.
    /// </summary>
    IXlsxCell TargetCell { get; set; }

    /// <summary>
    /// The cache of data for the pivot table. The pivot table is created
    /// from cached data, not up-to-date data in a worksheet.
    /// </summary>
    IXlsxPivotCache PivotCache { get; set; }

    bool MergeAndCenterWithLabels { get; set; } // MergeItem
    int RowLabelIndent { get; set; } // Indent

    /// <summary>
    /// Filter fields layout setting that indicates layout order of filter fields. The layout
    /// uses <see cref="FilterFieldsPageWrap"/> to determine when to break to a new row or
    /// column. Default value is <see cref="XlsxFilterAreaOrder.DownThenOver"/>.
    /// </summary>
    XlsxFilterAreaOrder FilterAreaOrder { get; set; }

    /// <summary>
    /// Specifies the number of page fields to display before starting another row or column.
    /// Value = 0 means unlimited.
    /// </summary>
    /// <exception cref="ArgumentOutOfRangeException">If value &lt; 0.</exception>
    int FilterFieldsPageWrap { get; set; } // PageWrap
    string ErrorValueReplacement { get; set; } // ErrorCaption
    string EmptyCellReplacement { get; set; } // MissingCaption
    bool AutofitColumns { get; set; } //UseAutoFormatting
    bool PreserveCellFormatting { get; set; } // PreserveFormatting

    bool ShowGrandTotalsRows { get; set; } // RowGrandTotals
    bool ShowGrandTotalsColumns { get; set; } // ColumnGrandTotals
    bool FilteredItemsInSubtotals { get; set; } // Subtotal filtered page items
    bool AllowMultipleFilters { get; set; } // MultipleFieldFilters
    bool UseCustomListsForSorting { get; set; } // CustomListSort

    bool ShowExpandCollapseButtons { get; set; }
    bool ShowContextualTooltips { get; set; }
    bool ShowPropertiesInTooltips { get; set; }
    bool DisplayCaptionsAndDropdowns { get; set; }
    bool ClassicPivotTableLayout { get; set; }
    bool ShowValuesRow { get; set; }
    bool ShowEmptyItemsOnRows { get; set; }
    bool ShowEmptyItemsOnColumns { get; set; }
    bool DisplayItemLabels { get; set; }
    bool SortFieldsAtoZ { get; set; }

    bool PrintExpandCollapsedButtons { get; set; }
    bool RepeatRowLabels { get; set; }
    bool PrintTitles { get; set; }

    bool EnableShowDetails { get; set; }
    bool EnableCellEditing { get; set; }

    IXlsxPivotTable CopyTo(IXlsxCell targetCell);

    IXlsxPivotTable SetName(string value);

    IXlsxPivotTable SetTitle(string value);

    IXlsxPivotTable SetDescription(string value);

    IXlsxPivotTable SetMergeAndCenterWithLabels(); IXlsxPivotTable SetMergeAndCenterWithLabels(bool value);

    IXlsxPivotTable SetRowLabelIndent(int value);

    IXlsxPivotTable SetFilterAreaOrder(XlsxFilterAreaOrder value);

    IXlsxPivotTable SetFilterFieldsPageWrap(int value);

    IXlsxPivotTable SetErrorValueReplacement(string value);

    IXlsxPivotTable SetEmptyCellReplacement(string value);

    IXlsxPivotTable SetAutofitColumns(); IXlsxPivotTable SetAutofitColumns(bool value);

    IXlsxPivotTable SetPreserveCellFormatting(); IXlsxPivotTable SetPreserveCellFormatting(bool value);

    IXlsxPivotTable SetShowGrandTotalsRows(); IXlsxPivotTable SetShowGrandTotalsRows(bool value);

    /// <summary>
    /// Should pivot table display a grand total for each row in the last column of a pivot
    /// table (it will enlarge pivot table for extra column).
    /// </summary>
    /// <remarks>
    /// This API has inverse row/column names than the Excel. Excel: <em>On for rows
    /// </em> should use this method <em>ShowGrandTotalsColumns</em>.
    /// </remarks>
    IXlsxPivotTable SetShowGrandTotalsColumns(); IXlsxPivotTable SetShowGrandTotalsColumns(bool value);

    IXlsxPivotTable SetFilteredItemsInSubtotals(); IXlsxPivotTable SetFilteredItemsInSubtotals(bool value);

    IXlsxPivotTable SetAllowMultipleFilters(); IXlsxPivotTable SetAllowMultipleFilters(bool value);

    IXlsxPivotTable SetUseCustomListsForSorting(); IXlsxPivotTable SetUseCustomListsForSorting(bool value);

    IXlsxPivotTable SetShowExpandCollapseButtons(); IXlsxPivotTable SetShowExpandCollapseButtons(bool value);

    IXlsxPivotTable SetShowContextualTooltips(); IXlsxPivotTable SetShowContextualTooltips(bool value);

    IXlsxPivotTable SetShowPropertiesInTooltips(); IXlsxPivotTable SetShowPropertiesInTooltips(bool value);

    IXlsxPivotTable SetDisplayCaptionsAndDropdowns(); IXlsxPivotTable SetDisplayCaptionsAndDropdowns(bool value);

    IXlsxPivotTable SetClassicPivotTableLayout(); IXlsxPivotTable SetClassicPivotTableLayout(bool value);

    IXlsxPivotTable SetShowValuesRow(); IXlsxPivotTable SetShowValuesRow(bool value);

    IXlsxPivotTable SetShowEmptyItemsOnRows(); IXlsxPivotTable SetShowEmptyItemsOnRows(bool value);

    IXlsxPivotTable SetShowEmptyItemsOnColumns(); IXlsxPivotTable SetShowEmptyItemsOnColumns(bool value);

    IXlsxPivotTable SetDisplayItemLabels(); IXlsxPivotTable SetDisplayItemLabels(bool value);

    IXlsxPivotTable SetSortFieldsAtoZ(); IXlsxPivotTable SetSortFieldsAtoZ(bool value);

    IXlsxPivotTable SetPrintExpandCollapsedButtons(); IXlsxPivotTable SetPrintExpandCollapsedButtons(bool value);

    IXlsxPivotTable SetRepeatRowLabels(); IXlsxPivotTable SetRepeatRowLabels(bool value);

    IXlsxPivotTable SetPrintTitles(); IXlsxPivotTable SetPrintTitles(bool value);


    IXlsxPivotTable SetEnableShowDetails(); IXlsxPivotTable SetEnableShowDetails(bool value);



    IXlsxPivotTable SetEnableCellEditing(); IXlsxPivotTable SetEnableCellEditing(bool value);

    IXlsxPivotTable SetColumnHeaderCaption(string value);

    IXlsxPivotTable SetRowHeaderCaption(string value);

    bool ShowRowHeaders { get; set; }
    bool ShowColumnHeaders { get; set; }
    bool ShowRowStripes { get; set; }
    bool ShowColumnStripes { get; set; }
    XlsxPivotSubtotals Subtotals { get; set; }

    /// <summary>
    /// Set the layout of the pivot table. It also changes layout of all pivot fields.
    /// </summary>
    XlsxPivotLayout Layout { set; }
    bool InsertBlankLines { set; }

    IXlsxPivotTable SetShowRowHeaders(); IXlsxPivotTable SetShowRowHeaders(bool value);

    IXlsxPivotTable SetShowColumnHeaders(); IXlsxPivotTable SetShowColumnHeaders(bool value);

    IXlsxPivotTable SetShowRowStripes(); IXlsxPivotTable SetShowRowStripes(bool value);

    IXlsxPivotTable SetShowColumnStripes(); IXlsxPivotTable SetShowColumnStripes(bool value);

    IXlsxPivotTable SetSubtotals(XlsxPivotSubtotals value);

    IXlsxPivotTable SetLayout(XlsxPivotLayout value);

    IXlsxPivotTable SetInsertBlankLines(); IXlsxPivotTable SetInsertBlankLines(bool value);

    IXlsxWorksheet Worksheet { get; }

    IXlsxPivotTableStyleFormats StyleFormats { get; }
}
