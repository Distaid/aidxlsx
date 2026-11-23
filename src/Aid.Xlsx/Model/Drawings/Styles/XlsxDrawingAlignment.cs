#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingAlignment(IXlsxDrawingStyle style) : IXlsxDrawingAlignment
{
    private readonly IXlsxDrawingStyle _style = style;

    public XlsxDrawingHorizontalAlignment Horizontal { get; set; }
    public IXlsxDrawingStyle SetHorizontal(XlsxDrawingHorizontalAlignment value) { Horizontal = value; return _style; }
    public XlsxDrawingVerticalAlignment Vertical { get; set; }
    public IXlsxDrawingStyle SetVertical(XlsxDrawingVerticalAlignment value) { Vertical = value; return _style; }
    public bool AutomaticSize { get; set; }
    public IXlsxDrawingStyle SetAutomaticSize() { AutomaticSize = true; return _style; }
    public IXlsxDrawingStyle SetAutomaticSize(bool value) { AutomaticSize = value; return _style; }
    public XlsxDrawingTextDirection Direction { get; set; }
    public IXlsxDrawingStyle SetDirection(XlsxDrawingTextDirection value) { Direction = value; return _style; }
    public XlsxDrawingTextOrientation Orientation { get; set; }
    public IXlsxDrawingStyle SetOrientation(XlsxDrawingTextOrientation value) { Orientation = value; return _style; }

}
