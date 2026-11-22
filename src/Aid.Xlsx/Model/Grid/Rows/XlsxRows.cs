using System.Collections;
using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Rows
{
    internal class XlsxRows : IXlsxRows
    {
        private readonly List<XlsxRow> _rowsCollection = new List<XlsxRow>();
        private readonly XlsxWorkbook _workbook;
        private readonly XlsxWorksheet? _worksheet;
        private readonly XlsxWorksheet? _defaultStyleSheet;

        /// <summary>
        /// This object represents all rows of the worksheet, even non-materialized ones.
        /// </summary>
        [MemberNotNullWhen(true, nameof(_worksheet))]
        private bool AllRowsOfSheet => _worksheet is not null;

        private bool IsMaterialized => _lazyEnumerable == null;

        private IEnumerable<XlsxRow>? _lazyEnumerable;
        private IEnumerable<XlsxRow> Rows => _lazyEnumerable ?? _rowsCollection.AsEnumerable();

        /// <summary>
        /// Create a new instance of <see cref="XlsxRows"/>.
        /// </summary>
        /// <param name="workbook">Workbook of the rows.</param>
        /// <param name="worksheet">If worksheet is specified it means that the created instance represents
        /// all rows on a worksheet so changing its height will affect all rows.</param>
        /// <param name="defaultStyleSheet">A sheet with a default style to use when initializing child entries.</param>
        /// <param name="lazyEnumerable">A predefined enumerator of <see cref="XlsxRow"/> to support lazy initialization.</param>
        public XlsxRows(XlsxWorkbook workbook, XlsxWorksheet? worksheet, XlsxWorksheet? defaultStyleSheet = null, IEnumerable<XlsxRow>? lazyEnumerable = null)
        {
            _workbook = workbook;
            _worksheet = worksheet;
            _defaultStyleSheet = defaultStyleSheet;
            _lazyEnumerable = lazyEnumerable;
        }

        #region IXlsxRows Members

        public IEnumerator<IXlsxRow> GetEnumerator()
        {
            return Rows.Cast<IXlsxRow>().OrderBy(r => r.RowNumber()).GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public double Height
        {
            set
            {
                Rows.ForEach(c => c.Height = value);
                if (!AllRowsOfSheet)
                    return;

                _worksheet.RowHeight = value;
                _worksheet.Internals.RowsCollection.ForEach(r => r.Value.Height = value);
            }
        }

        public void Delete()
        {
            if (AllRowsOfSheet)
            {
                _worksheet.Internals.RowsCollection.Clear();
                _worksheet.Internals.CellsCollection.Clear();
            }
            else
            {
                var toDelete = new Dictionary<IXlsxWorksheet, List<Int32>>();
                foreach (XlsxRow r in Rows)
                {
                    if (!toDelete.TryGetValue(r.Worksheet, out List<Int32>? list))
                    {
                        list = new List<Int32>();
                        toDelete.Add(r.Worksheet, list);
                    }

                    list.Add(r.RowNumber());
                }

                foreach (KeyValuePair<IXlsxWorksheet, List<int>> kp in toDelete)
                {
                    foreach (int r in kp.Value.OrderByDescending(r => r))
                        kp.Key.Row(r).Delete();
                }
            }
        }

        public IXlsxRows AdjustToContents()
        {
            Rows.ForEach(r => r.AdjustToContents());
            return this;
        }

        public IXlsxRows AdjustToContents(Int32 startColumn)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn));
            return this;
        }

        public IXlsxRows AdjustToContents(Int32 startColumn, Int32 endColumn)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn, endColumn));
            return this;
        }

        public IXlsxRows AdjustToContents(Double minHeight, Double maxHeight)
        {
            Rows.ForEach(r => r.AdjustToContents(minHeight, maxHeight));
            return this;
        }

        public IXlsxRows AdjustToContents(Int32 startColumn, Double minHeight, Double maxHeight)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn, minHeight, maxHeight));
            return this;
        }

        public IXlsxRows AdjustToContents(Int32 startColumn, Int32 endColumn, Double minHeight, Double maxHeight)
        {
            Rows.ForEach(r => r.AdjustToContents(startColumn, endColumn, minHeight, maxHeight));
            return this;
        }

        public void Hide()
        {
            Rows.ForEach(r => r.Hide());
        }

        public void Unhide()
        {
            Rows.ForEach(r => r.Unhide());
        }

        public void Group()
        {
            Group(false);
        }

        public void Group(Int32 outlineLevel)
        {
            Group(outlineLevel, false);
        }

        public void Ungroup()
        {
            Ungroup(false);
        }

        public void Group(Boolean collapse)
        {
            Rows.ForEach(r => r.Group(collapse));
        }

        public void Group(Int32 outlineLevel, Boolean collapse)
        {
            Rows.ForEach(r => r.Group(outlineLevel, collapse));
        }

        public void Ungroup(Boolean ungroupFromAll)
        {
            Rows.ForEach(r => r.Ungroup(ungroupFromAll));
        }

        public void Collapse()
        {
            Rows.ForEach(r => r.Collapse());
        }

        public void Expand()
        {
            Rows.ForEach(r => r.Expand());
        }

        public IXlsxCells Cells()
        {
            var cells = new XlsxCells(_workbook, false, XlsxCellsUsedOptions.AllContents);
            foreach (XlsxRow container in Rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXlsxCells CellsUsed()
        {
            var cells = new XlsxCells(_workbook, true, XlsxCellsUsedOptions.AllContents);
            foreach (XlsxRow container in Rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXlsxCells CellsUsed(XlsxCellsUsedOptions options)
        {
            var cells = new XlsxCells(_workbook, true, options);
            foreach (XlsxRow container in Rows)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXlsxRows AddHorizontalPageBreaks()
        {
            foreach (XlsxRow row in Rows)
                row.Worksheet.PageSetup.AddHorizontalPageBreak(row.RowNumber());
            return this;
        }

        public IXlsxStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }

        internal XlsxCellFormat Format
        {
            get
            {
                if (AllRowsOfSheet)
                {
                    return XlsxCellFormat.ForWorksheet(_worksheet);
                }

                return XlsxCellFormat.ForRows(_workbook, _defaultStyleSheet, Rows);

            }
        }

        #endregion IXlsxRows Members

        public void Add(XlsxRow row)
        {
            Materialize();
            _rowsCollection.Add(row);
        }

        public IXlsxRows Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
        {
            Rows.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }

        private void Materialize()
        {
            if (IsMaterialized)
                return;

            _rowsCollection.AddRange(Rows);
            _lazyEnumerable = null;
        }
    }
}
