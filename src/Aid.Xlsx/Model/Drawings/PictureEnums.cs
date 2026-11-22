#nullable disable

namespace Aid.Xlsx.Model.Drawings;

public enum XlsxMarkerPosition
{
    TopLeft,
    BottomRight
}

public enum XlsxPictureFormat
{
    Unknown,
    Bmp,
    Gif,
    Png,
    Tiff,
    Icon,
    Pcx,
    Jpeg,
    Emf,
    Wmf,
    Webp
}

public enum XlsxPicturePlacement
{
    MoveAndSize = 0,
    Move = 1,
    FreeFloating = 2
}
