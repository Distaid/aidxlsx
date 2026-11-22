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

    String Name { get; set; }
    String Title { get; set; }
    String Description { get; set; }

    String ColumnHeaderCaption { get; set; }
    String RowHeaderCaption { get; set; }

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

    Boolean MergeAndCenterWithLabels { get; set; } // MergeItem
    Int32 RowLabelIndent { get; set; } // Indent

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
    Int32 FilterFieldsPageWrap { get; set; } // PageWrap
    String ErrorValueReplacement { get; set; } // ErrorCaption
    String EmptyCellReplacement { get; set; } // MissingCaption
    Boolean AutofitColumns { get; set; } //UseAutoFormatting
    Boolean PreserveCellFormatting { get; set; } // PreserveFormatting

    Boolean ShowGrandTotalsRows { get; set; } // RowGrandTotals
    Boolean ShowGrandTotalsColumns { get; set; } // ColumnGrandTotals
    Boolean FilteredItemsInSubtotals { get; set; } // Subtotal filtered page items
    Boolean AllowMultipleFilters { get; set; } // MultipleFieldFilters
    Boolean UseCustomListsForSorting { get; set; } // CustomListSort

    Boolean ShowExpandCollapseButtons { get; set; }
    Boolean ShowContextualTooltips { get; set; }
    Boolean ShowPropertiesInTooltips { get; set; }
    Boolean DisplayCaptionsAndDropdowns { get; set; }
    Boolean ClassicPivotTableLayout { get; set; }
    Boolean ShowValuesRow { get; set; }
    Boolean ShowEmptyItemsOnRows { get; set; }
    Boolean ShowEmptyItemsOnColumns { get; set; }
    Boolean DisplayItemLabels { get; set; }
    Boolean SortFieldsAtoZ { get; set; }

    Boolean PrintExpandCollapsedButtons { get; set; }
    Boolean RepeatRowLabels { get; set; }
    Boolean PrintTitles { get; set; }

    Boolean EnableShowDetails { get; set; }
    Boolean EnableCellEditing { get; set; }

    IXlsxPivotTable CopyTo(IXlsxCell targetCell);

    IXlsxPivotTable SetName(String value);

    IXlsxPivotTable SetTitle(String value);

    IXlsxPivotTable SetDescription(String value);

    IXlsxPivotTable SetMergeAndCenterWithLabels(); IXlsxPivotTable SetMergeAndCenterWithLabels(Boolean value);

    IXlsxPivotTable SetRowLabelIndent(Int32 value);

    IXlsxPivotTable SetFilterAreaOrder(XlsxFilterAreaOrder value);

    IXlsxPivotTable SetFilterFieldsPageWrap(Int32 value);

    IXlsxPivotTable SetErrorValueReplacement(String value);

    IXlsxPivotTable SetEmptyCellReplacement(String value);

    IXlsxPivotTable SetAutofitColumns(); IXlsxPivotTable SetAutofitColumns(Boolean value);

    IXlsxPivotTable SetPreserveCellFormatting(); IXlsxPivotTable SetPreserveCellFormatting(Boolean value);

    IXlsxPivotTable SetShowGrandTotalsRows(); IXlsxPivotTable SetShowGrandTotalsRows(Boolean value);

    /// <summary>
    /// Should pivot table display a grand total for each row in the last column of a pivot
    /// table (it will enlarge pivot table for extra column).
    /// </summary>
    /// <remarks>
    /// This API has inverse row/column names than the Excel. Excel: <em>On for rows
    /// </em> should use this method <em>ShowGrandTotalsColumns</em>.
    /// </remarks>
    IXlsxPivotTable SetShowGrandTotalsColumns(); IXlsxPivotTable SetShowGrandTotalsColumns(Boolean value);

    IXlsxPivotTable SetFilteredItemsInSubtotals(); IXlsxPivotTable SetFilteredItemsInSubtotals(Boolean value);

    IXlsxPivotTable SetAllowMultipleFilters(); IXlsxPivotTable SetAllowMultipleFilters(Boolean value);

    IXlsxPivotTable SetUseCustomListsForSorting(); IXlsxPivotTable SetUseCustomListsForSorting(Boolean value);

    IXlsxPivotTable SetShowExpandCollapseButtons(); IXlsxPivotTable SetShowExpandCollapseButtons(Boolean value);

    IXlsxPivotTable SetShowContextualTooltips(); IXlsxPivotTable SetShowContextualTooltips(Boolean value);

    IXlsxPivotTable SetShowPropertiesInTooltips(); IXlsxPivotTable SetShowPropertiesInTooltips(Boolean value);

    IXlsxPivotTable SetDisplayCaptionsAndDropdowns(); IXlsxPivotTable SetDisplayCaptionsAndDropdowns(Boolean value);

    IXlsxPivotTable SetClassicPivotTableLayout(); IXlsxPivotTable SetClassicPivotTableLayout(Boolean value);

    IXlsxPivotTable SetShowValuesRow(); IXlsxPivotTable SetShowValuesRow(Boolean value);

    IXlsxPivotTable SetShowEmptyItemsOnRows(); IXlsxPivotTable SetShowEmptyItemsOnRows(Boolean value);

    IXlsxPivotTable SetShowEmptyItemsOnColumns(); IXlsxPivotTable SetShowEmptyItemsOnColumns(Boolean value);

    IXlsxPivotTable SetDisplayItemLabels(); IXlsxPivotTable SetDisplayItemLabels(Boolean value);

    IXlsxPivotTable SetSortFieldsAtoZ(); IXlsxPivotTable SetSortFieldsAtoZ(Boolean value);

    IXlsxPivotTable SetPrintExpandCollapsedButtons(); IXlsxPivotTable SetPrintExpandCollapsedButtons(Boolean value);

    IXlsxPivotTable SetRepeatRowLabels(); IXlsxPivotTable SetRepeatRowLabels(Boolean value);

    IXlsxPivotTable SetPrintTitles(); IXlsxPivotTable SetPrintTitles(Boolean value);


    IXlsxPivotTable SetEnableShowDetails(); IXlsxPivotTable SetEnableShowDetails(Boolean value);



    IXlsxPivotTable SetEnableCellEditing(); IXlsxPivotTable SetEnableCellEditing(Boolean value);

    IXlsxPivotTable SetColumnHeaderCaption(String value);

    IXlsxPivotTable SetRowHeaderCaption(String value);

    Boolean ShowRowHeaders { get; set; }
    Boolean ShowColumnHeaders { get; set; }
    Boolean ShowRowStripes { get; set; }
    Boolean ShowColumnStripes { get; set; }
    XlsxPivotSubtotals Subtotals { get; set; }

    /// <summary>
    /// Set the layout of the pivot table. It also changes layout of all pivot fields.
    /// </summary>
    XlsxPivotLayout Layout { set; }
    Boolean InsertBlankLines { set; }

    IXlsxPivotTable SetShowRowHeaders(); IXlsxPivotTable SetShowRowHeaders(Boolean value);

    IXlsxPivotTable SetShowColumnHeaders(); IXlsxPivotTable SetShowColumnHeaders(Boolean value);

    IXlsxPivotTable SetShowRowStripes(); IXlsxPivotTable SetShowRowStripes(Boolean value);

    IXlsxPivotTable SetShowColumnStripes(); IXlsxPivotTable SetShowColumnStripes(Boolean value);

    IXlsxPivotTable SetSubtotals(XlsxPivotSubtotals value);

    IXlsxPivotTable SetLayout(XlsxPivotLayout value);

    IXlsxPivotTable SetInsertBlankLines(); IXlsxPivotTable SetInsertBlankLines(Boolean value);

    IXlsxWorksheet Worksheet { get; }

    IXlsxPivotTableStyleFormats StyleFormats { get; }
}
