using System.Collections;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.DataValidation;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Ranges.Index;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Ranges
{
    internal class XlsxRanges : IXlsxRanges, IEnumerable<XlsxRange>
    {
        private readonly XlsxWorkbook _workbook;

        /// <summary>
        /// Normally, XlsxRanges collection includes ranges from a single worksheet, but not necessarily.
        /// </summary>
        private readonly Dictionary<IXlsxWorksheet, IXlsxRangeIndex<XlsxRange>> _indexes;
        private IEnumerable<XlsxRange> Ranges => _indexes.Values.SelectMany(index => index.GetAll());

        private IXlsxRangeIndex<XlsxRange> GetRangeIndex(IXlsxWorksheet worksheet)
        {
            if (!_indexes.TryGetValue(worksheet, out IXlsxRangeIndex<XlsxRange>? rangeIndex))
            {
                rangeIndex = new XlsxRangeIndex<XlsxRange>(worksheet);
                _indexes.Add(worksheet, rangeIndex);
            }

            return rangeIndex;
        }

        public XlsxRanges(XlsxWorksheet worksheet)
            : this(worksheet.Workbook)
        {
        }

        public XlsxRanges(XlsxWorkbook workbook)
        {
            _workbook = workbook;
            _indexes = new Dictionary<IXlsxWorksheet, IXlsxRangeIndex<XlsxRange>>();
        }

        internal XlsxCellFormat Format
        {
            get
            {
                var sheet = Ranges.FirstOrDefault()?.Worksheet;
                var areas = Ranges.Select(x => SheetArea.From(x.RangeAddress)).ToArray();
                return XlsxCellFormat.ForAreas(_workbook, areas, sheet);
            }
        }

        #region IXlsxRanges Members

        public IXlsxStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }

        IXlsxCells IXlsxRanges.Cells() => Cells();

        public IXlsxRanges Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
        {
            Ranges.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Add(XlsxRange range)
        {
            if (GetRangeIndex(range.Worksheet).Add(range))
                Count++;
        }

        public void Add(IXlsxRangeBase range)
        {
            Add((XlsxRange)range.AsRange());
        }

        public void Add(IXlsxCell cell)
        {
            Add(cell.AsRange());
        }

        public bool Remove(IXlsxRange range)
        {
            if (GetRangeIndex(range.Worksheet).Remove(range.RangeAddress))
            {
                Count--;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Removes ranges matching the criteria from the collection, optionally releasing their event handlers.
        /// </summary>
        /// <param name="match">Criteria to filter ranges. Only those ranges that satisfy the criteria will be removed.
        /// Null means the entire collection should be cleared.</param>
        /// <param name="releaseEventHandlers">Specify whether or not should removed ranges be unsubscribed from
        /// row/column shifting events. Until ranges are unsubscribed they cannot be collected by GC.</param>
        public void RemoveAll(Predicate<IXlsxRange>? match = null, bool releaseEventHandlers = true)
        {
            foreach (var index in _indexes.Values)
            {
                Count -= index.RemoveAll(match ?? (_ => true));
            }
        }

        public int Count { get; private set; }

        public IEnumerator<XlsxRange> GetEnumerator()
        {
            return Ranges
                .OrderBy(r => r.Worksheet.Position)
                .ThenBy(r => r.RangeAddress.FirstAddress.RowNumber)
                .ThenBy(r => r.RangeAddress.FirstAddress.ColumnNumber)
                .GetEnumerator();
        }

        IEnumerator<IXlsxRange> IEnumerable<IXlsxRange>.GetEnumerator()
        {
            return GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public Boolean Contains(IXlsxCell cell)
        {
            return GetIntersectedRanges((XlsxAddress)cell.Address).Any();
        }

        public Boolean Contains(IXlsxRange range)
        {
            return GetIntersectedRanges((XlsxRangeAddress)range.RangeAddress)
                .Any(r => r.Contains(range));
        }

        /// <summary>
        /// Filter ranges from a collection that intersect the specified address. Is much more efficient
        /// that using Linq expression .Where().
        /// </summary>
        public IEnumerable<IXlsxRange> GetIntersectedRanges(IXlsxRangeAddress rangeAddress)
        {
            var xlRangeAddress = (XlsxRangeAddress)rangeAddress;
            return GetIntersectedRanges(in xlRangeAddress);
        }

        internal IEnumerable<IXlsxRange> GetIntersectedRanges(in XlsxRangeAddress rangeAddress)
        {
            return GetRangeIndex(rangeAddress.Worksheet!)
                .GetIntersectedRanges(rangeAddress);
        }

        /// <summary>
        /// Filter ranges from a collection that intersect the specified address. Is much more efficient
        /// that using Linq expression .Where().
        /// </summary>
        public IEnumerable<IXlsxRange> GetIntersectedRanges(IXlsxAddress address)
        {
            var xlAddress = (XlsxAddress)address;
            return GetIntersectedRanges(in xlAddress);
        }

        internal IEnumerable<IXlsxRange> GetIntersectedRanges(in XlsxAddress address)
        {
            return GetRangeIndex(address.Worksheet)
                .GetIntersectedRanges(address);
        }

        public IEnumerable<IXlsxRange> GetIntersectedRanges(IXlsxCell cell)
        {
            return GetIntersectedRanges(cell.Address);
        }

        public IEnumerable<IXlsxDataValidation> DataValidation
        {
            get { return Ranges.Select(range => range.GetDataValidation()).Where(dv => dv != null); }
        }

        public IXlsxRanges AddToNamed(String rangeName)
        {
            return AddToNamed(rangeName, XlsxScope.Workbook);
        }

        public IXlsxRanges AddToNamed(String rangeName, XlsxScope scope)
        {
            return AddToNamed(rangeName, XlsxScope.Workbook, null);
        }

        public IXlsxRanges AddToNamed(String rangeName, XlsxScope scope, String? comment)
        {
            Ranges.ForEach(r => r.AddToNamed(rangeName, scope, comment));
            return this;
        }

        public XlsxCellValue Value
        {
            set { Ranges.ForEach(r => r.Value = value); }
        }

        public IXlsxRanges SetValue(XlsxCellValue value)
        {
            Ranges.ForEach(r => r.SetValue(value));
            return this;
        }

        public XlsxCells Cells()
        {
            var cells = new XlsxCells(_workbook, false, XlsxCellsUsedOptions.AllContents);
            foreach (XlsxRange container in Ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXlsxCells CellsUsed()
        {
            var cells = new XlsxCells(_workbook, true, XlsxCellsUsedOptions.AllContents);
            foreach (XlsxRange container in Ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXlsxCells CellsUsed(XlsxCellsUsedOptions options)
        {
            var cells = new XlsxCells(_workbook, true, options);
            foreach (XlsxRange container in Ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        #endregion IXlsxRanges Members

        public override string ToString()
        {
            String retVal = Ranges.Aggregate(String.Empty, (agg, r) => agg + (r.ToString() + ","));
            if (retVal.Length > 0) retVal = retVal.Substring(0, retVal.Length - 1);
            return retVal;
        }

        public override bool Equals(object? obj)
        {
            return Equals(obj as XlsxRanges);
        }

        public bool Equals(XlsxRanges? other)
        {
            if (other == null)
                return false;

            return Ranges.Count() == other.Ranges.Count() &&
                   Ranges.Select(thisRange => Enumerable.Contains(other.Ranges, thisRange)).All(foundOne => foundOne);
        }

        public override int GetHashCode()
        {
            return Ranges.Aggregate(0, (current, r) => current ^ r.GetHashCode());
        }

        public IXlsxDataValidation CreateDataValidation()
        {
            var firstRange = Ranges.First();
            var dataValidation = firstRange.Worksheet.DataValidations.Create(firstRange.SheetRange);
            foreach (var range in Ranges.Skip(1))
            {
                dataValidation.AddRange(range);
            }

            return dataValidation;
        }

        [Obsolete("Use CreateDataValidation() instead.")]
        public IXlsxDataValidation SetDataValidation()
        {
            return CreateDataValidation();
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }

        public IXlsxRanges Consolidate()
        {
            var engine = new XlsxRangeConsolidationEngine(_workbook, this);
            return engine.Consolidate();
        }
    }
}
