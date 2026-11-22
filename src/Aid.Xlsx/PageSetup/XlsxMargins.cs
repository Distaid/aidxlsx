#nullable disable

namespace Aid.Xlsx.PageSetup
{
    internal class XlsxMargins : IXlsxMargins
    {
        public Double Left { get; set; }
        public Double Right { get; set; }
        public Double Top { get; set; }
        public Double Bottom { get; set; }
        public Double Header { get; set; }
        public Double Footer { get; set; }

        public IXlsxMargins SetLeft(Double value) { Left = value; return this; }
        public IXlsxMargins SetRight(Double value) { Right = value; return this; }
        public IXlsxMargins SetTop(Double value) { Top = value; return this; }
        public IXlsxMargins SetBottom(Double value) { Bottom = value; return this; }
        public IXlsxMargins SetHeader(Double value) { Header = value; return this; }
        public IXlsxMargins SetFooter(Double value) { Footer = value; return this; }

    }
}
