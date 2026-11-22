using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Calculation;

/// <summary>
/// Context for <see cref="DependenciesVisitor"/>, it is used
/// to collect all objects a formula depends on during calculation.
/// </summary>
internal class DependenciesContext
{
    internal DependenciesContext(SheetArea formulaArea, XlsxWorkbook workbook)
    {
        FormulaArea = formulaArea;
        Workbook = workbook;
    }

    /// <summary>
    /// An area of a formula, in most cases just one cell, for array formulas area of cells.
    /// </summary>
    internal SheetArea FormulaArea { get; }

    internal XlsxWorkbook Workbook { get; }

    /// <summary>
    /// The result. Visitor adds all areas/names formula depends on to this.
    /// </summary>
    internal FormulaDependencies Dependencies { get; } = new();

    /// <summary>
    /// Add areas to a list of areas the formula depends on. Disregards duplicate entries.
    /// </summary>
    internal void AddAreas(List<SheetArea> sheetAreas) => Dependencies.AddAreas(sheetAreas);

    /// <summary>
    /// Add name to a list of names the formula depends on. Disregards duplicate entries.
    /// </summary>
    internal void AddName(XlsxName name) => Dependencies.AddName(name);
}
