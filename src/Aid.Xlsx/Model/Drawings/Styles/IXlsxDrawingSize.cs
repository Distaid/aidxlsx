#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

public interface IXlsxDrawingSize
{
    bool AutomaticSize { get; set; }
    double Height { get; set; }
    double Width { get; set; }

    IXlsxDrawingStyle SetAutomaticSize(); IXlsxDrawingStyle SetAutomaticSize(bool value);
    IXlsxDrawingStyle SetHeight(double value);
    IXlsxDrawingStyle SetWidth(double value);

}
