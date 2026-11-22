namespace Aid.Xlsx.Drawings.Styles
{
    internal class XlsxDrawingWeb : IXlsxDrawingWeb
    {
        private readonly IXlsxDrawingStyle _style;

        public XlsxDrawingWeb(IXlsxDrawingStyle style)
        {
            _style = style;
        }

        public String? AlternateText { get; set; }

        public IXlsxDrawingStyle SetAlternateText(String? value) { AlternateText = value; return _style; }
    }
}
