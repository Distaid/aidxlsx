#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Coordinates;

namespace Aid.Xlsx.Worksheets
{
    internal class XlsxSheetView : IXlsxSheetView
    {
        private XlsxAddress _topLeftCellAddress;
        private int _zoomScale;

        public XlsxSheetView(XlsxWorksheet worksheet)
        {
            Worksheet = worksheet;
            View = XlsxSheetViewOptions.Normal;

            ZoomScale = 100;
            ZoomScaleNormal = 100;
            ZoomScalePageLayoutView = 100;
            ZoomScaleSheetLayoutView = 100;
        }

        public XlsxSheetView(XlsxWorksheet worksheet, XlsxSheetView sheetView)
            : this(worksheet)
        {
            this.SplitRow = sheetView.SplitRow;
            this.SplitColumn = sheetView.SplitColumn;
            this.FreezePanes = sheetView.FreezePanes;
            this.TopLeftCellAddress = new XlsxAddress(this.Worksheet, sheetView.TopLeftCellAddress.RowNumber, sheetView.TopLeftCellAddress.ColumnNumber, sheetView.TopLeftCellAddress.FixedRow, sheetView.TopLeftCellAddress.FixedColumn);
        }

        public Boolean FreezePanes { get; set; }
        public Int32 SplitColumn { get; set; }
        public Int32 SplitRow { get; set; }

        IXlsxAddress IXlsxSheetView.TopLeftCellAddress { get => TopLeftCellAddress; set => TopLeftCellAddress = (XlsxAddress)value; }

        public XlsxAddress TopLeftCellAddress
        {
            get => _topLeftCellAddress;
            set
            {
                if (value.HasWorksheet && !value.Worksheet.Equals(this.Worksheet))
                    throw new ArgumentException($"The value should be on the same worksheet as the sheet view.");

                _topLeftCellAddress = value;
            }
        }

        public XlsxSheetViewOptions View { get; set; }

        IXlsxWorksheet IXlsxSheetView.Worksheet { get => Worksheet; }
        public XlsxWorksheet Worksheet { get; internal set; }

        public int ZoomScale
        {
            get { return _zoomScale; }
            set
            {
                _zoomScale = value;
                switch (View)
                {
                    case XlsxSheetViewOptions.Normal:
                        ZoomScaleNormal = value;
                        break;

                    case XlsxSheetViewOptions.PageBreakPreview:
                        ZoomScalePageLayoutView = value;
                        break;

                    case XlsxSheetViewOptions.PageLayout:
                        ZoomScaleSheetLayoutView = value;
                        break;
                }
            }
        }

        public int ZoomScaleNormal { get; set; }

        public int ZoomScalePageLayoutView { get; set; }

        public int ZoomScaleSheetLayoutView { get; set; }

        public void Freeze(Int32 rows, Int32 columns)
        {
            SplitRow = rows;
            SplitColumn = columns;
            FreezePanes = true;
        }

        public void FreezeColumns(Int32 columns)
        {
            SplitColumn = columns;
            FreezePanes = true;
        }

        public void FreezeRows(Int32 rows)
        {
            SplitRow = rows;
            FreezePanes = true;
        }

        public IXlsxSheetView SetView(XlsxSheetViewOptions value)
        {
            View = value;
            return this;
        }
    }
}
