#nullable disable

namespace Aid.Xlsx.Model.PageSetup;

internal class XlsxMargins : IXlsxMargins
{
    public double Left { get; set; }
    public double Right { get; set; }
    public double Top { get; set; }
    public double Bottom { get; set; }
    public double Header { get; set; }
    public double Footer { get; set; }

    public IXlsxMargins SetLeft(double value) { Left = value; return this; }
    public IXlsxMargins SetRight(double value) { Right = value; return this; }
    public IXlsxMargins SetTop(double value) { Top = value; return this; }
    public IXlsxMargins SetBottom(double value) { Bottom = value; return this; }
    public IXlsxMargins SetHeader(double value) { Header = value; return this; }
    public IXlsxMargins SetFooter(double value) { Footer = value; return this; }

}
