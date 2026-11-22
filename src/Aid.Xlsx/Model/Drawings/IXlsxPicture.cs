#nullable disable

// Keep this file CodeMaid organised and cleaned
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Drawings;

public interface IXlsxPicture
{
    IXlsxCell BottomRightCell { get; }

    /// <summary>
    /// Type of image. The supported formats are defined by OpenXML's ImagePartType.
    /// Default value is "jpeg"
    /// </summary>
    XlsxPictureFormat Format { get; }

    /// <summary>
    /// Current width of the picture in pixels.
    /// </summary>
    Int32 Width { get; set; }

    /// <summary>
    /// Current height of the picture in pixels.
    /// </summary>
    Int32 Height { get; set; }

    Int32 Id { get; }

    MemoryStream ImageStream { get; }

    Int32 Left { get; set; }

    /// <summary>
    /// Set the name of a picture.
    /// </summary>
    /// <exception cref="ArgumentException">Name is already used in the sheet, is null or empty.</exception>
    String Name { get; set; }

    /// <summary>
    /// Original height of the picture in pixels.
    /// </summary>
    Int32 OriginalHeight { get; }

    /// <summary>
    /// Original width of the picture in pixels.
    /// </summary>
    Int32 OriginalWidth { get; }

    XlsxPicturePlacement Placement { get; set; }

    Int32 Top { get; set; }

    IXlsxCell TopLeftCell { get; }

    IXlsxWorksheet Worksheet { get; }

    /// <summary>
    /// Create a copy of the picture on a different worksheet.
    /// </summary>
    /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
    /// <returns>A created copy of the picture.</returns>
    IXlsxPicture CopyTo(IXlsxWorksheet targetSheet);

    /// <summary>
    /// Deletes this picture.
    /// </summary>
    void Delete();

    /// <summary>
    /// Create a copy of the picture on the same worksheet.
    /// </summary>
    /// <returns>A created copy of the picture.</returns>
    IXlsxPicture Duplicate();

    System.Drawing.Point GetOffset(XlsxMarkerPosition position);

    IXlsxPicture MoveTo(Int32 left, Int32 top);

    IXlsxPicture MoveTo(IXlsxCell cell);

    IXlsxPicture MoveTo(IXlsxCell cell, Int32 xOffset, Int32 yOffset);

    IXlsxPicture MoveTo(IXlsxCell cell, System.Drawing.Point offset);

    IXlsxPicture MoveTo(IXlsxCell fromCell, IXlsxCell toCell);

    IXlsxPicture MoveTo(IXlsxCell fromCell, Int32 fromCellXOffset, Int32 fromCellYOffset, IXlsxCell toCell, Int32 toCellXOffset, Int32 toCellYOffset);

    IXlsxPicture MoveTo(IXlsxCell fromCell, System.Drawing.Point fromOffset, IXlsxCell toCell, System.Drawing.Point toOffset);

    IXlsxPicture Scale(Double factor, Boolean relativeToOriginal = false);

    IXlsxPicture ScaleHeight(Double factor, Boolean relativeToOriginal = false);

    IXlsxPicture ScaleWidth(Double factor, Boolean relativeToOriginal = false);

    IXlsxPicture WithPlacement(XlsxPicturePlacement value);

    IXlsxPicture WithSize(Int32 width, Int32 height);
}
