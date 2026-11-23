namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingWeb(IXlsxDrawingStyle style) : IXlsxDrawingWeb
{
    private readonly IXlsxDrawingStyle _style = style;

    public string? AlternateText { get; set; }

    public IXlsxDrawingStyle SetAlternateText(string? value) { AlternateText = value; return _style; }
}
