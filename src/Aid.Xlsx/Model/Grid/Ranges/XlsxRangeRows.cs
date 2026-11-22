#nullable disable

using System.Collections;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Ranges
{
    internal class XlsxRangeRows : IXlsxRangeRows
    {
        private readonly XlsxWorksheet _worksheet;
        private readonly List<XlsxRangeRow> _ranges = new List<XlsxRangeRow>();

        public XlsxRangeRows(XlsxWorksheet worksheet)
        {
            _worksheet = worksheet;
        }

        internal XlsxCellFormat Format
        {
            get
            {
                var areas = _ranges.Select(x => SheetArea.From(x.RangeAddress)).ToArray();
                return XlsxCellFormat.ForAreas(_worksheet.Workbook, areas, null);
            }
        }

        #region IXlsxRangeRows Members

        public IXlsxStyle Style
        {
            get => Format;
            set => Format.SetStyle(value);
        }

        public IXlsxRangeRows Clear(XlsxClearOptions clearOptions = XlsxClearOptions.All)
        {
            _ranges.ForEach(c => c.Clear(clearOptions));
            return this;
        }

        public void Delete()
        {
            _ranges.OrderByDescending(r => r.RowNumber()).ForEach(r => r.Delete());
            _ranges.Clear();
        }

        public void Add(IXlsxRangeRow range)
        {
            _ranges.Add((XlsxRangeRow)range);
        }

        public IEnumerator<IXlsxRangeRow> GetEnumerator()
        {
            return _ranges.Cast<IXlsxRangeRow>()
                          .OrderBy(r => r.Worksheet.Position)
                          .ThenBy(r => r.RowNumber())
                          .GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public IXlsxCells Cells()
        {
            var cells = new XlsxCells(_worksheet, false, XlsxCellsUsedOptions.AllContents);
            foreach (XlsxRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public IXlsxCells CellsUsed()
        {
            var cells = new XlsxCells(_worksheet, true, XlsxCellsUsedOptions.AllContents);
            foreach (XlsxRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }


        public IXlsxCells CellsUsed(XlsxCellsUsedOptions options)
        {
            var cells = new XlsxCells(_worksheet, true, options);
            foreach (XlsxRangeRow container in _ranges)
                cells.Add(container.RangeAddress);
            return cells;
        }

        public void Select()
        {
            foreach (var range in this)
                range.Select();
        }

        #endregion IXlsxRangeRows Members
    }
}
