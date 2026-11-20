#nullable disable

namespace Aid.Xlsx
{
    public interface IXLDrawingProperties
    {
        XLDrawingAnchor Positioning { get; set; }
        IXLDrawingStyle SetPositioning(XLDrawingAnchor value);

    }
}
