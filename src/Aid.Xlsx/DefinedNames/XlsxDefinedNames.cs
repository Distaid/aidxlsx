using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.CalcEngine.Visitors;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Formulas;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.DefinedNames
{
    /// <summary>
    /// A collection of a named ranges, either for workbook or for worksheet.
    /// </summary>
    internal class XlsxDefinedNames : IXlsxDefinedNames, IEnumerable<XlsxDefinedName>, ISheetListener
    {
        private readonly Dictionary<String, XlsxDefinedName> _namedRanges = new(XlsxHelper.NameComparer);

        internal XlsxWorkbook Workbook { get; }

        internal XlsxWorksheet? Worksheet { get; }

        internal XlsxNamedRangeScope Scope { get; }

        public XlsxDefinedNames(XlsxWorksheet worksheet)
            : this(worksheet.Workbook)
        {
            Worksheet = worksheet;
            Scope = XlsxNamedRangeScope.Worksheet;
        }

        public XlsxDefinedNames(XlsxWorkbook workbook)
        {
            Workbook = workbook;
            Scope = XlsxNamedRangeScope.Workbook;
        }

        #region IXlsxNamedRanges Members

        [Obsolete]
        IXlsxDefinedName IXlsxDefinedNames.NamedRange(String name) => DefinedName(name);

        IXlsxDefinedName IXlsxDefinedNames.DefinedName(String name) => DefinedName(name);

        internal XlsxDefinedName DefinedName(String name)
        {
            if (_namedRanges.TryGetValue(name, out XlsxDefinedName? range))
                return range;

            throw new KeyNotFoundException($"Name {name} not found.");
        }

        public IXlsxDefinedName Add(String name, String rangeAddress)
        {
            return Add(name, rangeAddress, null);
        }

        public IXlsxDefinedName Add(String name, IXlsxRange range)
        {
            return Add(name, range, null);
        }

        public IXlsxDefinedName Add(String name, IXlsxRanges ranges)
        {
            return Add(name, ranges, null);
        }

        public IXlsxDefinedName Add(String name, String rangeAddress, String? comment)
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
        internal IXlsxDefinedName Add(String name, String rangeAddress, String? comment, Boolean validateName, Boolean validateRangeAddress)
        {
            // When loading named ranges from an existing file, we do not validate the range address or name.
            if (validateRangeAddress)
            {
                var match = XlsxHelper.NamedRangeReferenceRegex.Match(rangeAddress);

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

                        if (Scope == XlsxNamedRangeScope.Workbook || !XlsxHelper.NamedRangeReferenceRegex.Match(range.ToString() ?? string.Empty).Success)
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

        public IXlsxDefinedName Add(String name, IXlsxRange range, String? comment)
        {
            var ranges = new XlsxRanges(Workbook) { range };
            return Add(name, ranges, comment);
        }

        public IXlsxDefinedName Add(String name, IXlsxRanges ranges, String? comment)
        {
            var formula = string.Join(",", ranges.Select(r => r.RangeAddress.ToStringFixed(XlsxReferenceStyle.A1, true)));
            var namedRange = new XlsxDefinedName(this, name, true, formula, comment);
            _namedRanges.Add(name, namedRange);
            return namedRange;
        }

        internal XlsxDefinedName Add(String name, XlsxDefinedName namedRange)
        {
            _namedRanges.Add(name, namedRange);
            return namedRange;
        }

        public void Delete(String rangeName)
        {
            _namedRanges.Remove(rangeName);
        }

        public void Delete(Int32 rangeIndex)
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

        public Boolean TryGetValue(String name, [NotNullWhen(true)] out IXlsxDefinedName? definedName)
        {
            if (TryGetScopedValue(name, out var sheetDefinedName))
            {
                definedName = sheetDefinedName;
                return true;
            }

            definedName = Scope == XlsxNamedRangeScope.Workbook
                ? Workbook.DefinedName(name)
                : null;

            return definedName is not null;
        }

        internal Boolean TryGetScopedValue(String name, [NotNullWhen(true)] out XlsxDefinedName? definedName)
        {
            if (_namedRanges.TryGetValue(name, out definedName))
            {
                return true;
            }

            return false;
        }

        public Boolean Contains(String name)
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
            foreach (var definedName in _namedRanges.Values)
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
}
