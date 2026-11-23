#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingMargins(IXlsxDrawingStyle style) : IXlsxDrawingMargins
{
    private readonly IXlsxDrawingStyle _style = style;

    public bool Automatic { get; set; }
    public IXlsxDrawingStyle SetAutomatic() { Automatic = true; return _style; }
    public IXlsxDrawingStyle SetAutomatic(bool value) { Automatic = value; return _style; }
    double _left;
    public double Left { get { return _left; } set { _left = value; Automatic = false; } }
    public IXlsxDrawingStyle SetLeft(double value) { Left = value; return _style; }
    double _right;
    public double Right { get { return _right; } set { _right = value; Automatic = false; } }
    public IXlsxDrawingStyle SetRight(double value) { Right = value; return _style; }
    double _top;
    public double Top { get { return _top; } set { _top = value; Automatic = false; } }
    public IXlsxDrawingStyle SetTop(double value) { Top = value; return _style; }
    double _bottom;
    public double Bottom { get { return _bottom; } set { _bottom = value; Automatic = false; } }
    public IXlsxDrawingStyle SetBottom(double value) { Bottom = value; return _style; }
    public double All
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
    public IXlsxDrawingStyle SetAll(double value) { All = value; return _style; }
}
