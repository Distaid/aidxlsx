#nullable disable

using Aid.Xlsx.Model.PivotTables.Areas;

namespace Aid.Xlsx.Model.PivotTables.PivotStyleFormats;

internal class XlsxPivotTableAxisFieldStyleFormats(XlsxPivotTable pivotTable, XlsxPivotTableAxisField axisField) : IXlsxPivotFieldStyleFormats
{
    private readonly XlsxPivotTable _pivotTable = pivotTable;
    private readonly XlsxPivotTableAxisField _axisField = axisField;

    #region IXlsxPivotFieldStyleFormats

    public IXlsxPivotValueStyleFormat DataValuesFormat => new XlsxPivotValueStyleFormat(_pivotTable, _axisField.Offset);

    public IXlsxPivotStyleFormat Header
    {
        get
        {
            /*
             * <x:pivotArea field="4"
             *  type="button"
             *  axis="axisCol"
             *  fieldPosition="0"/>
             *
             * The area must have field position and axis, otherwise the style is not correctly
             * displayed.
             */
            // If table is not compact, each field has it's own header and thus pivot area must
            // contain correct position of the field on the axis. If table is compact, there is
            // only one header and its position is first in axis, because it's the only one.
            var fieldPosition = _pivotTable.Compact ? 0 : _axisField.Position;
            var fieldAxis = _axisField.Axis;
            var headerArea = new XlsxPivotArea
            {
                Field = _axisField.Offset,
                Type = XlsxPivotAreaType.Button,
                Axis = fieldAxis,
                FieldPosition = (uint)fieldPosition,
            };

            return new XlsxPivotStyleFormat(
                _pivotTable,
                area => XlsxPivotAreaComparer.Instance.Equals(area, headerArea),
                () => headerArea);
        }
    }

    public IXlsxPivotStyleFormat Label
    {
        get
        {
            /* <x:pivotArea type="normal"
			 *              dataOnly="0"
			 *              labelOnly="1">
			 *     <x:references count="1">
			 *         <x:reference field="4"/>
			 *	   </x:references>
			 * </x:pivotArea>
             */
            var labelArea = new XlsxPivotArea
            {
                DataOnly = false,
                LabelOnly = true,
            };
            labelArea.AddReference(new XlsxPivotReference
            {
                Field = (uint)_axisField.Offset,
            });

            return new XlsxPivotStyleFormat(
                _pivotTable,
                area => XlsxPivotAreaComparer.Instance.Equals(area, labelArea),
                () => labelArea);
        }
    }

    public IXlsxPivotStyleFormat Subtotal
    {
        get
        {
            /* <pivotArea outline="0">
             *   <references count="1">
             *     <reference field="0"
             *                count="0"
             *                defaultSubtotal="1"/>
             *   </references>
             * </pivotArea>
             */
            // Subtotal fields in reference can't mix default and custom subtotals. It always must
            // reference only one type. Excel doesn't select correct area if they are mixed.
            // The outline flag has weird behavior, but is required for subtotals of last field in
            // an axis with multiple fields (i.e. subtotals are displayed at the bottom).
            var subtotals = _axisField.Subtotals.ToHashSet();
            var subtotalArea = new XlsxPivotArea
            {
                Outline = false
            };
            subtotalArea.AddReference(new XlsxPivotReference
            {
                Field = unchecked((uint)_axisField.Offset),
                Subtotals = subtotals,
            });

            return new XlsxPivotStyleFormat(
                _pivotTable,
                area => XlsxPivotAreaComparer.Instance.Equals(area, subtotalArea),
                () => subtotalArea);
        }
    }

    #endregion
}
