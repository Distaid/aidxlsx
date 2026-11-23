namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingWeb(IXlsxDrawingStyle style) : IXlsxDrawingWeb
{
    private readonly IXlsxDrawingStyle _style = style;

    public String? AlternateText { get; set; }

    public IXlsxDrawingStyle SetAlternateText(String? value) { AlternateText = value; return _style; }
}
