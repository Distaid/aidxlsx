#nullable disable

namespace Aid.Xlsx.Drawings.Styles
{
    public enum XlsxDrawingTextDirection
    {
        Context,
        LeftToRight,
        RightToLeft
    }
    public enum XlsxDrawingTextOrientation
    {
        LeftToRight,
        Vertical,
        BottomToTop,
        TopToBottom
    }
    public enum XlsxDrawingHorizontalAlignment { Left, Justify, Center, Right, Distributed }
    public enum XlsxDrawingVerticalAlignment { Top, Justify, Center, Bottom, Distributed }
    public interface IXlsxDrawingAlignment
    {
        XlsxDrawingHorizontalAlignment Horizontal { get; set; }
        XlsxDrawingVerticalAlignment Vertical { get; set; }
        Boolean AutomaticSize { get; set; }
        XlsxDrawingTextDirection Direction { get; set; }
        XlsxDrawingTextOrientation Orientation { get; set; }

        IXlsxDrawingStyle SetHorizontal(XlsxDrawingHorizontalAlignment value);
        IXlsxDrawingStyle SetVertical(XlsxDrawingVerticalAlignment value);
        IXlsxDrawingStyle SetAutomaticSize(); IXlsxDrawingStyle SetAutomaticSize(Boolean value);
        IXlsxDrawingStyle SetDirection(XlsxDrawingTextDirection value);
        IXlsxDrawingStyle SetOrientation(XlsxDrawingTextOrientation value);

    }
}
