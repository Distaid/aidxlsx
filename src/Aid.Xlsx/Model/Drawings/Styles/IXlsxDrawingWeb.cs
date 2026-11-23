namespace Aid.Xlsx.Model.Drawings.Styles;

public interface IXlsxDrawingWeb
{
    string? AlternateText { get; set; }
    IXlsxDrawingStyle SetAlternateText(string? value);
}
