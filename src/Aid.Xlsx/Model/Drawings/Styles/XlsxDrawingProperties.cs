#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingProperties(IXlsxDrawingStyle style) : IXlsxDrawingProperties
{
    private readonly IXlsxDrawingStyle _style = style;

    public XlsxDrawingAnchor Positioning { get; set; }
    public IXlsxDrawingStyle SetPositioning(XlsxDrawingAnchor value) { Positioning = value; return _style; }

}
