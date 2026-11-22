#nullable disable

using Aid.Xlsx.Drawings.Styles;

namespace Aid.Xlsx.Drawings
{
    public enum XlsxDrawingAnchor { MoveAndSizeWithCells, MoveWithCells, Absolute }
    public interface IXlsxDrawing<T>
    {
        Int32 ShapeId { get; }

        Boolean Visible { get; set; }
        T SetVisible();
        T SetVisible(Boolean hidden);

        ////String Name { get; set; }
        ////T SetName(String name);

        ////String Description { get; set; }
        ////T SetDescription(String description);

        IXlsxDrawingPosition Position { get; }

        Int32 ZOrder { get; set; }
        T SetZOrder(Int32 zOrder);

        //Boolean HorizontalFlip { get; set; }
        //T SetHorizontalFlip();
        //T SetHorizontalFlip(Boolean horizontalFlip);

        //Boolean VerticalFlip { get; set; }
        //T SetVerticalFlip();
        //T SetVerticalFlip(Boolean verticalFlip);

        //Int32 Rotation { get; set; }
        //T SetRotation(Int32 rotation);

        //Int32 ExtentLength { get; set; }
        //T SetExtentLength(Int32 ExtentLength);

        //Int32 ExtentWidth { get; set; }
        //T SetExtentWidth(Int32 extentWidth);

        IXlsxDrawingStyle Style { get; }

    }
}
