#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

public interface IXlsxDrawingSize
{
    Boolean AutomaticSize { get; set; }
    Double Height { get; set; }
    Double Width { get; set; }

    IXlsxDrawingStyle SetAutomaticSize(); IXlsxDrawingStyle SetAutomaticSize(Boolean value);
    IXlsxDrawingStyle SetHeight(Double value);
    IXlsxDrawingStyle SetWidth(Double value);

}
