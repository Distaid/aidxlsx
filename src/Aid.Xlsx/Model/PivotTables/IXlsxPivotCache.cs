using Aid.Xlsx.Model.Exceptions;

namespace Aid.Xlsx.Model.PivotTables;

/// <summary>
/// A cache of pivot data - essentially a collection of fields and their values that can be
/// displayed by a <see cref="IXlsxPivotTable"/>. Data for the cache are retrieved from
/// an area (a table or a range). The pivot cache data are <strong>cached</strong>, i.e.
/// the data in the source are not immediately updated once the data in a worksheet change.
/// </summary>
public interface IXlsxPivotCache
{
    /// <summary>
    /// Get names of all fields in the source, in left to right order. Every field name is unique.
    /// </summary>
    /// <remarks>
    /// The field names are case insensitive. The field names of the cached
    /// source might differ from actual names of the columns
    /// in the data cells.
    /// </remarks>
    IReadOnlyList<string> FieldNames { get; }

    /// <summary>
    /// Gets the number of unused items in shared items to allow before discarding unused items.
    /// </summary>
    /// <remarks>
    /// Shared items are distinct values of a source field values. Updating them can be expensive
    /// and this controls, when should the cache be updated. Application-dependent attribute.
    /// </remarks>
    /// <value>Default value is <see cref="XlsxItemsToRetain.Automatic"/>.</value>
    XlsxItemsToRetain ItemsToRetainPerField { get; set; }

    /// <summary>
    /// Will Excel refresh the cache when it opens the workbook.
    /// </summary>
    /// <value>Default value is <c>false</c>.</value>
    bool RefreshDataOnOpen { get; set; }

    /// <summary>
    /// Should the cached values of the pivot source be saved into the workbook file?
    /// If source data are not saved, they will have to be refreshed from the source
    /// reference which might cause a change in the table values.
    /// </summary>
    /// <value>Default value is <c>true</c>.</value>
    bool SaveSourceData { get; set; }

    /// <summary>
    /// Refresh data in the pivot source from the source reference data.
    /// </summary>
    /// <exception cref="InvalidReferenceException">The data source for the pivot table can't be found.</exception>
    IXlsxPivotCache Refresh();

    /// <inheritdoc cref="ItemsToRetainPerField"/>
    IXlsxPivotCache SetItemsToRetainPerField(XlsxItemsToRetain value);

    /// <inheritdoc cref="RefreshDataOnOpen"/>
    /// <remarks>Sets the value to <c>true</c>.</remarks>
    IXlsxPivotCache SetRefreshDataOnOpen();

    /// <inheritdoc cref="RefreshDataOnOpen"/>
    IXlsxPivotCache SetRefreshDataOnOpen(bool value);

    /// <inheritdoc cref="SaveSourceData"/>
    /// <remarks>Sets the value to <c>true</c>.</remarks>
    IXlsxPivotCache SetSaveSourceData();

    /// <inheritdoc cref="SaveSourceData"/>
    IXlsxPivotCache SetSaveSourceData(bool value);
}
