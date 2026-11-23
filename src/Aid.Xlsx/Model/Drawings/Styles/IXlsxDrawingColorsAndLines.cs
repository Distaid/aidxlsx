#nullable disable

using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Drawings.Styles;

public enum XlsxDashStyle
{
    Solid,
    RoundDot,
    SquareDot,
    Dash,
    DashDot,
    LongDash,
    LongDashDot,
    LongDashDotDot
}
public enum XlsxLineStyle
{
    Single, ThinThin, ThinThick, ThickThin, ThickBetweenThin
}
public interface IXlsxDrawingColorsAndLines
{
    XlsxColor FillColor { get; set; }
    double FillTransparency { get; set; }
    XlsxColor LineColor { get; set; }
    double LineTransparency { get; set; }
    double LineWeight { get; set; }
    XlsxDashStyle LineDash { get; set; }
    XlsxLineStyle LineStyle { get; set; }

    IXlsxDrawingStyle SetFillColor(XlsxColor value);
    IXlsxDrawingStyle SetFillTransparency(double value);
    IXlsxDrawingStyle SetLineColor(XlsxColor value);
    IXlsxDrawingStyle SetLineTransparency(double value);
    IXlsxDrawingStyle SetLineWeight(double value);
    IXlsxDrawingStyle SetLineDash(XlsxDashStyle value);
    IXlsxDrawingStyle SetLineStyle(XlsxLineStyle value);

}
