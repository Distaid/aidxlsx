#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingSize(IXlsxDrawingStyle style) : IXlsxDrawingSize
{
    private readonly IXlsxDrawingStyle _style = style;

    public Boolean AutomaticSize { get { return _style.Alignment.AutomaticSize; } set { _style.Alignment.AutomaticSize = value; } }
    public IXlsxDrawingStyle SetAutomaticSize() { AutomaticSize = true; return _style; }
    public IXlsxDrawingStyle SetAutomaticSize(Boolean value) { AutomaticSize = value; return _style; }
    public Double Height { get; set; }
    public IXlsxDrawingStyle SetHeight(Double value) { Height = value; return _style; }
    public Double Width { get; set; }
    public IXlsxDrawingStyle SetWidth(Double value) { Width = value; return _style; }
}
