#nullable disable

using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingColorsAndLines : IXlsxDrawingColorsAndLines
{
    private readonly IXlsxDrawingStyle _style;

    public XlsxDrawingColorsAndLines(IXlsxDrawingStyle style)
    {
        _style = style;
    }
    public XlsxColor FillColor { get; set; }
    public IXlsxDrawingStyle SetFillColor(XlsxColor value) { FillColor = value; return _style; }
    public Double FillTransparency { get; set; }
    public IXlsxDrawingStyle SetFillTransparency(Double value) { FillTransparency = value; return _style; }
    public XlsxColor LineColor { get; set; }
    public IXlsxDrawingStyle SetLineColor(XlsxColor value) { LineColor = value; return _style; }
    public Double LineTransparency { get; set; }
    public IXlsxDrawingStyle SetLineTransparency(Double value) { LineTransparency = value; return _style; }
    public Double LineWeight { get; set; }
    public IXlsxDrawingStyle SetLineWeight(Double value) { LineWeight = value; return _style; }
    public XlsxDashStyle LineDash { get; set; }
    public IXlsxDrawingStyle SetLineDash(XlsxDashStyle value) { LineDash = value; return _style; }
    public XlsxLineStyle LineStyle { get; set; }
    public IXlsxDrawingStyle SetLineStyle(XlsxLineStyle value) { LineStyle = value; return _style; }

}
