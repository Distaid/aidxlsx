#nullable disable

using Aid.Xlsx.Styles.Colors;

namespace Aid.Xlsx.Drawings.Styles
{
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
        Double FillTransparency { get; set; }
        XlsxColor LineColor { get; set; }
        Double LineTransparency { get; set; }
        Double LineWeight { get; set; }
        XlsxDashStyle LineDash { get; set; }
        XlsxLineStyle LineStyle { get; set; }

        IXlsxDrawingStyle SetFillColor(XlsxColor value);
        IXlsxDrawingStyle SetFillTransparency(Double value);
        IXlsxDrawingStyle SetLineColor(XlsxColor value);
        IXlsxDrawingStyle SetLineTransparency(Double value);
        IXlsxDrawingStyle SetLineWeight(Double value);
        IXlsxDrawingStyle SetLineDash(XlsxDashStyle value);
        IXlsxDrawingStyle SetLineStyle(XlsxLineStyle value);

    }
}
