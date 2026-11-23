using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using Aid.Xlsx.Calculation.Visitors;
using Aid.Xlsx.Formulas.References.Transforms;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.DefinedNames;

/// <summary>
/// A collection of a named ranges, either for workbook or for worksheet.
/// </summary>
internal class XlsxDefinedNames(XlsxWorkbook workbook) : IXlsxDefinedNames, IEnumerable<XlsxDefinedName>, ISheetListener
{
    private readonly Dictionary<string, XlsxDefinedName> _namedRanges = [with(XlsxHelper.NameComparer)];

    internal XlsxWorkbook Workbook { get; } = workbook;

    internal XlsxWorksheet? Worksheet { get; }

    internal XlsxNamedRangeScope Scope { get; } = XlsxNamedRangeScope.Workbook;

    public XlsxDefinedNames(XlsxWorksheet worksheet)
        : this(worksheet.Workbook)
    {
        Worksheet = worksheet;
        Scope = XlsxNamedRangeScope.Worksheet;
    }

    #region IXlsxNamedRanges Members

    [Obsolete]
    IXlsxDefinedName IXlsxDefinedNames.NamedRange(string name) => DefinedName(name);

    IXlsxDefinedName IXlsxDefinedNames.DefinedName(string name) => DefinedName(name);

    internal XlsxDefinedName DefinedName(string name)
    {
        if (_namedRanges.TryGetValue(name, out XlsxDefinedName? range))
            return range;

        throw new KeyNotFoundException($"Name {name} not found.");
    }

    public IXlsxDefinedName Add(string name, string rangeAddress)
    {
        return Add(name, rangeAddress, null);
    }

    public IXlsxDefinedName Add(string name, IXlsxRange range)
    {
        return Add(name, range, null);
    }

    public IXlsxDefinedName Add(string name, IXlsxRanges ranges)
    {
        return Add(name, ranges, null);
    }

    public IXlsxDefinedName Add(string name, string rangeAddress, string? comment)
    {
        return Add(name, rangeAddress, comment, validateName: true, validateRangeAddress: true);
    }

    /// <summary>
    /// Adds the specified range name.
    /// </summary>
    /// <param name="name">Name of the range.</param>
    /// <param name="rangeAddress">The range address.</param>
    /// <param name="comment">The comment.</param>
    /// <param name="validateName">if set to <c>true</c> validates the name.</param>
    /// <param name="validateRangeAddress">if set to <c>true</c> range address will be checked for validity.</param>
    /// <exception cref="NotSupportedException"></exception>
    /// <exception cref="ArgumentException">
    /// For named ranges in the workbook scope, specify the sheet name in the reference.
    /// </exception>
    internal IXlsxDefinedName Add(string name, string rangeAddress, string? comment, bool validateName, bool validateRangeAddress)
    {
        // When loading named ranges from an existing file, we do not validate the range address or name.
        if (validateRangeAddress)
        {
            Match match = XlsxHelper.NamedRangeReferenceRegex.Match(rangeAddress);

            if (!match.Success)
            {
                if (XlsxHelper.IsValidRangeAddress(rangeAddress))
                {
                    IXlsxRange? range;
                    if (Scope == XlsxNamedRangeScope.Worksheet)
                        range = Worksheet!.Range(rangeAddress);
                    else if (Scope == XlsxNamedRangeScope.Workbook)
                        range = Workbook.Range(rangeAddress);
                    else
                        throw new NotSupportedException($"Scope {Scope} is not supported");

                    if (range == null)
                        throw new ArgumentException(string.Format(
                            "The range address '{0}' for the named range '{1}' is not a valid range.", rangeAddress,
                            name));

                    if (Scope == XlsxNamedRangeScope.Workbook || !XlsxHelper.NamedRangeReferenceRegex.IsMatch(range.ToString() ?? string.Empty))
                        throw new ArgumentException(
                            "For named ranges in the workbook scope, specify the sheet name in the reference.");

                    rangeAddress = range.ToString() ?? string.Empty;
                }
            }
        }

        var namedRange = new XlsxDefinedName(this, name, validateName, rangeAddress, comment);
        _namedRanges.Add(name, namedRange);
        return namedRange;
    }

    public IXlsxDefinedName Add(string name, IXlsxRange range, string? comment)
    {
        var ranges = new XlsxRanges(Workbook) { range };
        return Add(name, ranges, comment);
    }

    public IXlsxDefinedName Add(string name, IXlsxRanges ranges, string? comment)
    {
        var formula = string.Join(",", ranges.Select(r => r.RangeAddress.ToStringFixed(XlsxReferenceStyle.A1, true)));
        var namedRange = new XlsxDefinedName(this, name, true, formula, comment);
        _namedRanges.Add(name, namedRange);
        return namedRange;
    }

    internal XlsxDefinedName Add(string name, XlsxDefinedName namedRange)
    {
        _namedRanges.Add(name, namedRange);
        return namedRange;
    }

    public void Delete(string rangeName)
    {
        _namedRanges.Remove(rangeName);
    }

    public void Delete(int rangeIndex)
    {
        _namedRanges.Remove(_namedRanges.ElementAt(rangeIndex).Key);
    }

    public void DeleteAll()
    {
        _namedRanges.Clear();
    }

    /// <summary>
    /// Returns a subset of named ranges that do not have invalid references.
    /// </summary>
    public IEnumerable<IXlsxDefinedName> ValidNamedRanges()
    {
        return _namedRanges.Values.Where(nr => nr.IsValid);
    }

    /// <summary>
    /// Returns a subset of named ranges that do have invalid references.
    /// </summary>
    public IEnumerable<IXlsxDefinedName> InvalidNamedRanges()
    {
        return _namedRanges.Values.Where(nr => !nr.IsValid);
    }

    #endregion IXlsxNamedRanges Members

    IEnumerator<XlsxDefinedName> IEnumerable<XlsxDefinedName>.GetEnumerator() => GetEnumerator();

    IEnumerator<IXlsxDefinedName> IEnumerable<IXlsxDefinedName>.GetEnumerator() => GetEnumerator();

    public Dictionary<string, XlsxDefinedName>.ValueCollection.Enumerator GetEnumerator()
    {
        return _namedRanges.Values.GetEnumerator();
    }

    #region IEnumerable Members

    System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    #endregion IEnumerable Members

    public bool TryGetValue(string name, [NotNullWhen(true)] out IXlsxDefinedName? definedName)
    {
        if (TryGetScopedValue(name, out XlsxDefinedName? sheetDefinedName))
        {
            definedName = sheetDefinedName;
            return true;
        }

        definedName = Scope == XlsxNamedRangeScope.Workbook
            ? Workbook.DefinedName(name)
            : null;

        return definedName is not null;
    }

    internal bool TryGetScopedValue(string name, [NotNullWhen(true)] out XlsxDefinedName? definedName)
    {
        if (_namedRanges.TryGetValue(name, out definedName))
        {
            return true;
        }

        return false;
    }

    public bool Contains(string name)
    {
        if (_namedRanges.ContainsKey(name)) return true;

        if (Scope == XlsxNamedRangeScope.Workbook)
            return Workbook.DefinedName(name) is not null;
        else
            return false;
    }

    internal void OnWorksheetDeleted(string worksheetName)
    {
        _namedRanges.Values
            .ForEach(nr => nr.OnWorksheetDeleted(worksheetName));
    }

    #region ISheetListner

    void ISheetListener.OnInsertAreaAndShiftDown(XlsxWorksheet sheet, Area insertedArea)
    {
        var insertedBookArea = new SheetArea(sheet.Name, insertedArea);
        var refMod = new ReferenceShiftOnInsertRefModVisitor(insertedBookArea, true);
        ShiftReferences(refMod);
    }

    void ISheetListener.OnInsertAreaAndShiftRight(XlsxWorksheet sheet, Area insertedArea)
    {
        var insertedBookArea = new SheetArea(sheet.Name, insertedArea);
        var refMod = new ReferenceShiftOnInsertRefModVisitor(insertedBookArea, false);
        ShiftReferences(refMod);
    }

    void ISheetListener.OnDeleteAreaAndShiftLeft(XlsxWorksheet sheet, Area deletedArea)
    {
        var deletedBookArea = new SheetArea(sheet.Name, deletedArea);
        var refMod = new ReferenceShiftOnDeleteRefModVisitor(deletedBookArea, XlsxShiftDeletedCells.ShiftCellsLeft);
        ShiftReferences(refMod);
    }

    void ISheetListener.OnDeleteAreaAndShiftUp(XlsxWorksheet sheet, Area deletedArea)
    {
        var deletedBookArea = new SheetArea(sheet.Name, deletedArea);
        var refMod = new ReferenceShiftOnDeleteRefModVisitor(deletedBookArea, XlsxShiftDeletedCells.ShiftCellsUp);
        ShiftReferences(refMod);
    }

    private void ShiftReferences(CopyVisitor refMod)
    {
        foreach (XlsxDefinedName definedName in _namedRanges.Values)
        {
            var nameFormula = definedName.RefersTo;

            // Defined name formula should never rely on a context info. Formula should contain
            // only absolute references with a sheet -> use empty sheet that should never match
            // and the cell point is thus never used and can be left at A1.
            var shiftedFormula = FormulaConverter.ModifyA1(nameFormula, string.Empty, 1, 1, refMod);
            definedName.SetRefersTo(shiftedFormula);
        }
    }
    #endregion
}
