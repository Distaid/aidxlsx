namespace Aid.Xlsx.PivotTables.Areas;

/// <summary>
/// A representation of a single row/column axis values in a <see cref="XlsxPivotTableAxis"/>.
/// </summary>
/// <remarks>
/// Represents 18.10.1.44 i (Row Items) and 18.10.1.96 x (Member Property Index).
/// </remarks>
internal class XlsxPivotFieldAxisItem
{
    public XlsxPivotFieldAxisItem(XlsxPivotItemType itemType, int dataItem, IEnumerable<int> fieldItems)
    {
        ItemType = itemType;
        DataItem = dataItem;
        FieldItem = fieldItems.ToList();
    }

    /// <summary>
    /// Each item is an index to field items of corresponding field from
    /// <see cref="XlsxPivotTableAxis.Fields"/>. Value <c>1048832</c> specifies that no item appears
    /// at the position. 
    /// </summary>
    internal List<int> FieldItem { get; }

    /// <summary>
    /// Type of item.
    /// </summary>
    internal XlsxPivotItemType ItemType { get; }

    /// <summary>
    /// If this item (row/column) contains 'data' field, this contains an index into the <see cref="XlsxPivotTable.DataFields"/>
    /// that should be used as a value. The value for 'data' field in the <see cref="FieldItem"/> is ignored, but Excel fills
    /// same number as this index.
    /// </summary>
    internal int DataItem { get; }
}
