#nullable disable

using Aid.Xlsx.Model.Drawings.Styles;

namespace Aid.Xlsx.Model.Drawings;

internal class XlsxDrawing<T> : IXlsxDrawing<T>
{
    internal T Container;
    public XlsxDrawing()
    {
        Style = new XlsxDrawingStyle();
        Position = new XlsxDrawingPosition();
    }

    public int ShapeId { get; internal set; }

    public bool Visible { get; set; }
    public T SetVisible()
    {
        Visible = true;
        return Container;
    }
    public T SetVisible(bool hidden)
    {
        Visible = hidden;
        return Container;
    }

    public string Name { get; set; }
    public T SetName(string name)
    {
        Name = name;
        return Container;
    }

    public string Description { get; set; }
    public T SetDescription(string description)
    {
        Description = description;
        return Container;
    }

    public IXlsxDrawingPosition Position { get; private set; }

    public int ZOrder { get; set; }
    public T SetZOrder(int zOrder)
    {
        ZOrder = zOrder;
        return Container;
    }

    public bool HorizontalFlip { get; set; }
    public T SetHorizontalFlip()
    {
        HorizontalFlip = true;
        return Container;
    }
    public T SetHorizontalFlip(bool horizontalFlip)
    {
        HorizontalFlip = horizontalFlip;
        return Container;
    }

    public bool VerticalFlip { get; set; }
    public T SetVerticalFlip()
    {
        VerticalFlip = true;
        return Container;
    }
    public T SetVerticalFlip(bool verticalFlip)
    {
        VerticalFlip = verticalFlip;
        return Container;
    }

    public int Rotation { get; set; }
    public T SetRotation(int rotation)
    {
        Rotation = rotation;
        return Container;
    }

    public int OffsetX { get; set; }
    public T SetOffsetX(int offsetX)
    {
        OffsetX = offsetX;
        return Container;
    }

    public int OffsetY { get; set; }
    public T SetOffsetY(int offsetY)
    {
        OffsetY = offsetY;
        return Container;
    }

    public int ExtentLength { get; set; }
    public T SetExtentLength(int extentLength)
    {
        ExtentLength = extentLength;
        return Container;
    }

    public int ExtentWidth { get; set; }
    public T SetExtentWidth(int extentWidth)
    {
        ExtentWidth = extentWidth;
        return Container;
    }

    public IXlsxDrawingStyle Style { get; private set; }


}
