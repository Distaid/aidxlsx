using Aid.Xlsx.Formulas.References;
using Aid.Xlsx.Formulas.References.Transforms;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Calculation.Visitors;

/// <summary>
/// A RefModVisitor that adjusts a reference in a formula when an area is deleted and shifted up/left.
/// </summary>
internal class ReferenceShiftOnDeleteRefModVisitor : CopyVisitor
{
    private readonly SheetArea _deletedBookArea;
    private readonly XlsxShiftDeletedCells _shift;

    public ReferenceShiftOnDeleteRefModVisitor(SheetArea deletedBookArea, XlsxShiftDeletedCells shift)
    {
        _deletedBookArea = deletedBookArea;
        _shift = shift;
    }

    public override TransformedSymbol SheetReference(ModContext ctx, SymbolRange range, string sheet, ReferenceArea reference)
    {
        return ShiftFormulaReferences(ctx, range, sheet, reference);
    }

    public override TransformedSymbol Reference(ModContext ctx, SymbolRange range, ReferenceArea reference)
    {
        return ShiftFormulaReferences(ctx, range, null, reference);
    }

    private TransformedSymbol ShiftFormulaReferences(ModContext ctx, SymbolRange range, string? referenceSheetName, ReferenceArea referenceToShift)
    {
        if (!XlsxHelper.SheetComparer.Equals(_deletedBookArea.Name, referenceSheetName ?? ctx.Sheet))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var wouldSplitArea = _shift == XlsxShiftDeletedCells.ShiftCellsUp
            ? !referenceToShift.TryDeleteAndShiftUp(_deletedBookArea.Area, out var shiftedReference)
            : !referenceToShift.TryDeleteAndShiftLeft(_deletedBookArea.Area, out shiftedReference);

        // Return original reference if the shift would cause a split
        if (wouldSplitArea)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        // If reference was shifted out of sheet, return #REF!
        if (shiftedReference is null)
            return TransformedSymbol.ToText(ctx.Formula, range, XlsxHelper.RefError);

        // Do not allocate a new string unless necessary
        if (referenceToShift == shiftedReference.Value)
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var shiftedReferenceA1 = shiftedReference.Value.GetDisplayStringA1(referenceSheetName);
        return TransformedSymbol.ToText(ctx.Formula, range, shiftedReferenceA1);
    }
}
