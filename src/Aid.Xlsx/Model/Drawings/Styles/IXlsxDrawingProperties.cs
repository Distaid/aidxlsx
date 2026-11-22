#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

public interface IXlsxDrawingProperties
{
    XlsxDrawingAnchor Positioning { get; set; }
    IXlsxDrawingStyle SetPositioning(XlsxDrawingAnchor value);

}
