#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingMargins(IXlsxDrawingStyle style) : IXlsxDrawingMargins
{
    private readonly IXlsxDrawingStyle _style = style;

    public Boolean Automatic { get; set; }
    public IXlsxDrawingStyle SetAutomatic() { Automatic = true; return _style; }
    public IXlsxDrawingStyle SetAutomatic(Boolean value) { Automatic = value; return _style; }
    Double _left;
    public Double Left { get { return _left; } set { _left = value; Automatic = false; } }
    public IXlsxDrawingStyle SetLeft(Double value) { Left = value; return _style; }
    Double _right;
    public Double Right { get { return _right; } set { _right = value; Automatic = false; } }
    public IXlsxDrawingStyle SetRight(Double value) { Right = value; return _style; }
    Double _top;
    public Double Top { get { return _top; } set { _top = value; Automatic = false; } }
    public IXlsxDrawingStyle SetTop(Double value) { Top = value; return _style; }
    Double _bottom;
    public Double Bottom { get { return _bottom; } set { _bottom = value; Automatic = false; } }
    public IXlsxDrawingStyle SetBottom(Double value) { Bottom = value; return _style; }
    public Double All
    {
        set
        {
            _left = value;
            _right = value;
            _top = value;
            _bottom = value;
            Automatic = false;
        }
    }
    public IXlsxDrawingStyle SetAll(Double value) { All = value; return _style; }
}
