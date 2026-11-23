#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingSize(IXlsxDrawingStyle style) : IXlsxDrawingSize
{
    private readonly IXlsxDrawingStyle _style = style;

    public bool AutomaticSize { get { return _style.Alignment.AutomaticSize; } set { _style.Alignment.AutomaticSize = value; } }
    public IXlsxDrawingStyle SetAutomaticSize() { AutomaticSize = true; return _style; }
    public IXlsxDrawingStyle SetAutomaticSize(bool value) { AutomaticSize = value; return _style; }
    public double Height { get; set; }
    public IXlsxDrawingStyle SetHeight(double value) { Height = value; return _style; }
    public double Width { get; set; }
    public IXlsxDrawingStyle SetWidth(double value) { Width = value; return _style; }
}
