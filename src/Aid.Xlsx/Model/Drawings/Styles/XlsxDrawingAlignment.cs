#nullable disable

namespace Aid.Xlsx.Drawings.Styles
{
    internal class XlsxDrawingAlignment : IXlsxDrawingAlignment
    {
        private readonly IXlsxDrawingStyle _style;

        public XlsxDrawingAlignment(IXlsxDrawingStyle style)
        {
            _style = style;
        }
        public XlsxDrawingHorizontalAlignment Horizontal { get; set; }
        public IXlsxDrawingStyle SetHorizontal(XlsxDrawingHorizontalAlignment value) { Horizontal = value; return _style; }
        public XlsxDrawingVerticalAlignment Vertical { get; set; }
        public IXlsxDrawingStyle SetVertical(XlsxDrawingVerticalAlignment value) { Vertical = value; return _style; }
        public Boolean AutomaticSize { get; set; }
        public IXlsxDrawingStyle SetAutomaticSize() { AutomaticSize = true; return _style; }
        public IXlsxDrawingStyle SetAutomaticSize(Boolean value) { AutomaticSize = value; return _style; }
        public XlsxDrawingTextDirection Direction { get; set; }
        public IXlsxDrawingStyle SetDirection(XlsxDrawingTextDirection value) { Direction = value; return _style; }
        public XlsxDrawingTextOrientation Orientation { get; set; }
        public IXlsxDrawingStyle SetOrientation(XlsxDrawingTextOrientation value) { Orientation = value; return _style; }

    }
}
