#nullable disable

namespace Aid.Xlsx.Drawings.Styles
{
    public interface IXlsxDrawingProperties
    {
        XlsxDrawingAnchor Positioning { get; set; }
        IXlsxDrawingStyle SetPositioning(XlsxDrawingAnchor value);

    }
}
