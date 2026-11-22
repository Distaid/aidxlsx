#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingProperties : IXlsxDrawingProperties
{
    private readonly IXlsxDrawingStyle _style;

    public XlsxDrawingProperties(IXlsxDrawingStyle style)
    {
        _style = style;
    }
    public XlsxDrawingAnchor Positioning { get; set; }
    public IXlsxDrawingStyle SetPositioning(XlsxDrawingAnchor value) { Positioning = value; return _style; }

}
