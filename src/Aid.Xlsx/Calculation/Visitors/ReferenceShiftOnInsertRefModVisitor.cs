using Aid.Xlsx.Formulas.References;
using Aid.Xlsx.Formulas.References.Transforms;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Coordinates;

namespace Aid.Xlsx.Calculation.Visitors;

/// <summary>
/// A RefModVisitor that adjusts a reference in a formula when an area is inserted and cells are shifted down/right.
/// </summary>
internal class ReferenceShiftOnInsertRefModVisitor : CopyVisitor
{
    private readonly SheetArea _insertedBookArea;
    private readonly bool _shiftDown;

    internal ReferenceShiftOnInsertRefModVisitor(SheetArea insertedBookArea, bool shiftDown)
    {
        _insertedBookArea = insertedBookArea;
        _shiftDown = shiftDown;
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
        if (!XlsxHelper.SheetComparer.Equals(_insertedBookArea.Name, referenceSheetName ?? ctx.Sheet))
            return TransformedSymbol.CopyOriginal(ctx.Formula, range);

        var wouldSplitArea = _shiftDown
            ? !referenceToShift.TryInsertAndShiftDown(_insertedBookArea.Area, out ReferenceArea? shiftedReference)
            : !referenceToShift.TryInsertAndShiftRight(_insertedBookArea.Area, out shiftedReference);

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
