using System.Diagnostics;
using Aid.Xlsx.CalcEngine.Visitors;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Formulas;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Tables;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.DefinedNames;

[DebuggerDisplay("{_name}:{_formula}")]
internal class XlsxDefinedName : IXlsxDefinedName, IWorkbookListener
{
    private readonly XlsxDefinedNames _container;
    private String _name;
    private String _formula = null!;
    private FormulaReferences _references = null!;

    internal XlsxDefinedName(XlsxDefinedNames container, String name, Boolean validateName, String formula, String? comment)
    {
        // Excel accepts invalid names per grammar (e.g. `[Foo]Bar`) as a valid name and they can
        // encountered in existing workbooks. We shouldn't throw exception on load.
        if (validateName)
        {
            if (!XlsxHelper.ValidateName("named range", name, out var error))
                throw new ArgumentException(error, nameof(name));
        }

        _container = container;
        _name = name;
        RefersTo = formula;
        Visible = true;
        Comment = comment;
    }

    public bool IsValid => !_references.ContainsRefError;

    public String Name
    {
        get => _name;
        set
        {
            if (XlsxHelper.NameComparer.Equals(_name, value))
                return;

            if (!XlsxHelper.ValidateName("named range", value, out var error))
                throw new ArgumentException(error, nameof(value));

            if (_container.Contains(value))
                throw new InvalidOperationException($"There is already a name '{value}'.");

            _container.Delete(_name);
            _name = value;
            _container.Add(_name, this);
        }
    }

    public IXlsxRanges Ranges => _references.GetExternalRanges(_container.Workbook, new Point(1, 1));

    public String? Comment { get; set; }

    public Boolean Visible { get; set; }

    public XlsxNamedRangeScope Scope => _container.Scope;

    public String RefersTo
    {
        get => _formula;
        set
        {
            if (value is null)
                throw new ArgumentNullException();

            var formula = value.TrimFormulaEqual();
            if (string.IsNullOrWhiteSpace(formula))
                throw new ArgumentException("Formula can't be empty.");

            var references = FormulaReferences.ForFormula(formula);
            if (references.References.Any())
            {
                // `[MS-XLSX] 2.2.2.5: The formula MUST NOT use the local-cell-reference production
                // rule.` Excel will refuse to load a workbook with such a defined name (e.g. `A1`).
                // In theory, defined name should support bang references as a replacement for local
                // references, but ClosedParser doesn't support it yet.
                throw new ArgumentException($"Formula '{formula}' contains references without a sheet.");
            }

            _references = references;
            _formula = formula;
        }
    }

    IXlsxDefinedName IXlsxDefinedName.CopyTo(IXlsxWorksheet targetSheet) => CopyTo((XlsxWorksheet)targetSheet);

    void IXlsxDefinedName.Delete() => _container.Delete(Name);

    /// <summary>
    /// Get sheet references found in the formula in A1. Doesn't return tables or name references,
    /// only what has col/row coordinates.
    /// </summary>
    internal IReadOnlyList<String> SheetReferencesList => _references.SheetReferences.Select(x => x.GetA1()).ToList();

    internal XlsxDefinedName CopyTo(XlsxWorksheet targetSheet)
    {
        var sheet = _container.Worksheet;
        if (targetSheet == sheet)
            throw new InvalidOperationException("Cannot copy named range to the worksheet it already belongs to.");

        if (sheet is null)
            throw new InvalidOperationException("Cannot copy workbook scoped defined name.");

        var targetTables = targetSheet.Tables.ToDictionary<XlsxTable, Area>(x => x.SheetRange);
        var tableRenames = new Dictionary<string, string>();
        foreach (var table in sheet.Tables)
        {
            if (targetTables.TryGetValue(table.SheetRange, out var targetTable))
            {
                tableRenames.Add(table.Name, targetTable.Name);
            }
        }

        var copiedFormula = FormulaConverter.ModifyA1(_formula, sheet.Name, 1, 1, new RenameRefModVisitor
        {
            Sheets = new Dictionary<string, string?> { { sheet.Name, targetSheet.Name } },
            Tables = tableRenames,
        });
        var copiedName = new XlsxDefinedName(targetSheet.DefinedNames, Name, false, copiedFormula, Comment);
        return targetSheet.DefinedNames.Add(Name, copiedName);
    }

    public IXlsxDefinedName SetRefersTo(IXlsxRangeBase range)
    {
        return SetRefersTo(RangeToFixed(range));
    }

    public IXlsxDefinedName SetRefersTo(IXlsxRanges ranges)
    {
        var unionFormula = string.Join(",", ranges.Select(RangeToFixed));
        return SetRefersTo(unionFormula);
    }

    public IXlsxDefinedName SetRefersTo(String formula)
    {
        RefersTo = formula;
        return this;
    }

    public override string ToString()
    {
        return _formula;
    }

    internal void Add(String rangeAddress)
    {
        var byExclamation = rangeAddress.Split('!');
        var wsName = byExclamation[0].Replace("'", "");
        var rng = byExclamation[1];
        var rangeToAdd = _container.Workbook.WorksheetsInternal.Worksheet(wsName).Range(rng);

        var ranges = new XlsxRanges(_container.Workbook) { rangeToAdd };
        RefersTo = _formula + "," + string.Join(",", ranges.Select<XlsxRange, string>(RangeToFixed));
    }

    void IWorkbookListener.OnSheetRenamed(string oldSheetName, string newSheetName)
    {
        RenameFormulaSheet(oldSheetName, newSheetName);
    }

    internal void OnWorksheetDeleted(string worksheetName)
    {
        RenameFormulaSheet(worksheetName, null);
    }

    private void RenameFormulaSheet(string oldSheetName, string? newSheetName)
    {
        if (!_references.ContainsSheet(oldSheetName))
            return;

        var modified = FormulaConverter.ModifyA1(_formula, newSheetName ?? string.Empty, 1, 1, new RenameRefModVisitor
        {
            Sheets = new Dictionary<string, string?> { { oldSheetName, newSheetName } }
        });

        RefersTo = modified;
    }

    private static string RangeToFixed(IXlsxRangeBase range)
    {
        return range.RangeAddress.ToStringFixed(XlsxReferenceStyle.A1, true);
    }
}
