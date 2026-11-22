namespace Aid.Xlsx.Drawings.Styles
{
    public interface IXlsxDrawingWeb
    {
        String? AlternateText { get; set; }
        IXlsxDrawingStyle SetAlternateText(String? value);
    }
}
