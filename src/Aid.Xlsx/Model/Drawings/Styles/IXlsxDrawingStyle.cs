#nullable disable

namespace Aid.Xlsx.Drawings.Styles
{
    public interface IXlsxDrawingStyle
    {
        //IXlsxDrawingFont Font { get; }
        IXlsxDrawingAlignment Alignment { get; }
        IXlsxDrawingColorsAndLines ColorsAndLines { get; }
        IXlsxDrawingSize Size { get; }
        IXlsxDrawingProtection Protection { get; }
        IXlsxDrawingProperties Properties { get; }
        IXlsxDrawingMargins Margins { get; }
        IXlsxDrawingWeb Web { get; }
    }
}
