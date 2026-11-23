// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.PivotTables.Areas;
using Aid.Xlsx.Model.PivotTables.PivotValues;

namespace Aid.Xlsx.Model.PivotTables.PivotStyleFormats;

internal class XlsxPivotValueStyleFormat : XlsxPivotStyleFormatBase, IXlsxPivotValueStyleFormat
{
    /// <summary>
    /// A list of references that specify which data cells will be styled.
    /// A data cell will be styled, if it lies on all referenced fields.
    /// The term "lie on" means that either column or a row of data cell
    /// intersects a label cell of referenced field.
    /// </summary>
    private readonly List<FieldReference> _fieldReferences = [];

    public XlsxPivotValueStyleFormat(XlsxPivotTable pivotTable, FieldIndex fieldIndex)
        : base(pivotTable)
    {
        _fieldReferences.Add(new FieldReference(fieldIndex));
    }

    #region IXlsxPivotValueStyleFormat members

    public IXlsxPivotValueStyleFormat AndWith(IXlsxPivotField field)
    {
        _fieldReferences.Add(new FieldReference(field.Offset));
        return this;
    }

    public IXlsxPivotValueStyleFormat AndWith(IXlsxPivotField field, Predicate<XlsxCellValue>? predicate)
    {
        FieldIndex fieldIndex = field.Offset;
        if (fieldIndex.IsDataField)
            throw new ArgumentException("Field is a 'data' field.", nameof(field));

        if (predicate is null)
            return AndWith(field);

        XlsxPivotTableField pivotField = PivotTable.PivotFields[fieldIndex];
        var filteredItems = pivotField.GetAllItems(predicate)
            .WhereNotNull(fieldItem => fieldItem.ItemIndex)
            .Select(itemIndex => (uint)itemIndex)
            .ToList();

        _fieldReferences.Add(new FieldReference(field.Offset, filteredItems));
        return this;
    }

    public IXlsxPivotValueStyleFormat ForValueField(IXlsxPivotValue valueField)
    {
        var valuesIndex = PivotTable.DataFields.IndexOf(valueField);
        if (valuesIndex == -1)
            throw new ArgumentOutOfRangeException($"Field '{valueField.CustomName}' is not among value fields of the pivot table.");

        _fieldReferences.Add(new FieldReference(FieldIndex.DataField, new[] { (uint)valuesIndex }));
        return this;
    }

    #endregion IXlsxPivotValueStyleFormat members

    internal override XlsxPivotArea GetCurrentArea()
    {
        var area = new XlsxPivotArea();
        foreach (FieldReference fieldReference in _fieldReferences)
        {
            var reference = new XlsxPivotReference
            {
                Field = unchecked((uint?)fieldReference.FieldIndex.Value)
            };
            if (fieldReference.Items is not null)
            {
                foreach (var item in fieldReference.Items)
                    reference.AddFieldItem(item);
            }

            area.AddReference(reference);
        }

        return area;
    }

    internal override bool Filter(XlsxPivotArea area)
    {
        XlsxPivotArea currentArea = GetCurrentArea();
        return XlsxPivotAreaComparer.Instance.Equals(area, currentArea);
    }

    private record FieldReference(FieldIndex FieldIndex, IReadOnlyList<uint>? Items = null);
}
