namespace Aid.Xlsx.Model.Drawings.Styles;

public interface IXlsxDrawingWeb
{
    String? AlternateText { get; set; }
    IXlsxDrawingStyle SetAlternateText(String? value);
}
