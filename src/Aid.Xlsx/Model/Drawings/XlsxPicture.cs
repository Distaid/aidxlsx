#nullable disable

// Keep this file CodeMaid organised and cleaned
using System.Diagnostics;
using System.Drawing;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Rendering.Images;

namespace Aid.Xlsx.Model.Drawings;

[DebuggerDisplay("{Name}")]
internal sealed class XlsxPicture : IXlsxPicture, IDisposable
{
    private int _height;
    private int _id;
    private string _name = string.Empty;
    private int _width;
    private bool _disposed;

    internal XlsxPicture(XlsxWorksheet worksheet, Stream stream)
        : this(worksheet, stream, XlsxPictureFormat.Unknown)
    {
    }

    internal XlsxPicture(IXlsxWorksheet worksheet, Stream stream, XlsxPictureFormat format)
        : this(worksheet)
    {
        ArgumentNullException.ThrowIfNull(stream);

        XlsxPictureInfo info = worksheet.Workbook.GraphicEngine.GetPictureInfo(stream, format);
        Init(info);

        ImageStream = new MemoryStream();
        stream.Position = 0;
        stream.CopyTo(ImageStream);
        ImageStream.Seek(0, SeekOrigin.Begin);
    }

    private XlsxPicture(IXlsxWorksheet worksheet)
    {
        this.Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        this.Placement = XlsxPicturePlacement.MoveAndSize;
        this.Markers = new Dictionary<XlsxMarkerPosition, XlsxMarker>()
        {
            [XlsxMarkerPosition.TopLeft] = null,
            [XlsxMarkerPosition.BottomRight] = null
        };

        // Calculate default picture ID
        IEnumerable<IXlsxPicture> allPictures = worksheet.Workbook.Worksheets.SelectMany(ws => ws.Pictures);
        var freeId = allPictures.Select(x => x.Id).DefaultIfEmpty(0).Max() + 1;
        _id = freeId;
    }

    public IXlsxCell BottomRightCell
    {
        get => Markers[XlsxMarkerPosition.BottomRight].Cell;
        private set
        {
            if (!value.Worksheet.Equals(this.Worksheet))
                throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");

            Markers[XlsxMarkerPosition.BottomRight] = new XlsxMarker(value);
        }
    }

    public XlsxPictureFormat Format { get; private set; } = XlsxPictureFormat.Unknown;

    public int Height
    {
        get => _height;
        set
        {
            if (Placement == XlsxPicturePlacement.MoveAndSize)
                throw new ArgumentException("To set the height, the placement should be FreeFloating or Move");

            _height = value;
        }
    }

    public int Id
    {
        get => _id;
        internal set
        {
            if ((Worksheet.Pictures.FirstOrDefault(p => p.Id.Equals(value)) ?? this) != this)
                throw new ArgumentException($"The picture ID '{value}' already exists.");

            _id = value;
        }
    }

    public MemoryStream ImageStream
    {
        get
        {
            ThrowIfDisposed();
            return field;
        }
        private init;
    }

    public int Left
    {
        get => Markers[XlsxMarkerPosition.TopLeft]?.Offset.X ?? 0;
        set
        {
            if (Placement != XlsxPicturePlacement.FreeFloating)
                throw new ArgumentException("To set the left-hand offset, the placement should be FreeFloating");

            Markers[XlsxMarkerPosition.TopLeft] = new XlsxMarker(Worksheet.Cell(1, 1), new System.Drawing.Point(value, Top));
        }
    }

    public string Name
    {
        get => _name;
        set
        {
            if (_name == value)
                return;

            if ((Worksheet.Pictures.FirstOrDefault(p => p.Name.Equals(value, StringComparison.OrdinalIgnoreCase)) ?? this) != this)
                throw new ArgumentException($"The picture name '{value}' already exists.");

            SetName(value);
        }
    }

    public int OriginalHeight { get; private set; }

    public int OriginalWidth { get; private set; }

    public XlsxPicturePlacement Placement { get; set; }

    public int Top
    {
        get { return Markers[XlsxMarkerPosition.TopLeft]?.Offset.Y ?? 0; }
        set
        {
            if (this.Placement != XlsxPicturePlacement.FreeFloating)
                throw new ArgumentException("To set the top offset, the placement should be FreeFloating");

            Markers[XlsxMarkerPosition.TopLeft] = new XlsxMarker(Worksheet.Cell(1, 1), new System.Drawing.Point(this.Left, value));
        }
    }

    public IXlsxCell TopLeftCell
    {
        get
        {
            return Markers[XlsxMarkerPosition.TopLeft].Cell;
        }

        private set
        {
            if (!value.Worksheet.Equals(this.Worksheet))
                throw new InvalidOperationException("A picture and its anchor cells must be on the same worksheet");

            this.Markers[XlsxMarkerPosition.TopLeft] = new XlsxMarker(value);
        }
    }

    public int Width
    {
        get { return _width; }
        set
        {
            if (this.Placement == XlsxPicturePlacement.MoveAndSize)
                throw new ArgumentException("To set the width, the placement should be FreeFloating or Move");
            _width = value;
        }
    }

    public IXlsxWorksheet Worksheet { get; }

    internal IDictionary<XlsxMarkerPosition, XlsxMarker> Markers { get; private set; }

    internal string RelId { get; set; }

    /// <summary>
    /// Create a copy of the picture on a different worksheet.
    /// </summary>
    /// <param name="targetSheet">The worksheet to which the picture will be copied.</param>
    /// <returns>A created copy of the picture.</returns>
    public IXlsxPicture CopyTo(IXlsxWorksheet targetSheet)
    {
        return CopyTo((XlsxWorksheet)targetSheet);
    }

    public void Delete()
    {
        Worksheet.Pictures.Delete(this.Name);
    }

    /// <summary>
    /// Create a copy of the picture on the same worksheet.
    /// </summary>
    /// <returns>A created copy of the picture.</returns>
    public IXlsxPicture Duplicate()
    {
        return CopyTo(Worksheet);
    }

    public System.Drawing.Point GetOffset(XlsxMarkerPosition position)
    {
        return Markers[position].Offset;
    }

    public IXlsxPicture MoveTo(int left, int top)
    {
        this.Placement = XlsxPicturePlacement.FreeFloating;
        this.Left = left;
        this.Top = top;
        return this;
    }

    public IXlsxPicture MoveTo(IXlsxCell cell)
    {
        return MoveTo(cell, 0, 0);
    }

    public IXlsxPicture MoveTo(IXlsxCell cell, int xOffset, int yOffset)
    {
        return MoveTo(cell, new System.Drawing.Point(xOffset, yOffset));
    }

    public IXlsxPicture MoveTo(IXlsxCell cell, System.Drawing.Point offset)
    {
        this.Placement = XlsxPicturePlacement.Move;
        this.TopLeftCell = cell ?? throw new ArgumentNullException(nameof(cell));
        this.Markers[XlsxMarkerPosition.TopLeft].Offset = offset;
        return this;
    }

    public IXlsxPicture MoveTo(IXlsxCell fromCell, IXlsxCell toCell)
    {
        return MoveTo(fromCell, 0, 0, toCell, 0, 0);
    }

    public IXlsxPicture MoveTo(IXlsxCell fromCell, int fromCellXOffset, int fromCellYOffset, IXlsxCell toCell, int toCellXOffset, int toCellYOffset)
    {
        return MoveTo(fromCell, new System.Drawing.Point(fromCellXOffset, fromCellYOffset), toCell, new System.Drawing.Point(toCellXOffset, toCellYOffset));
    }

    public IXlsxPicture MoveTo(IXlsxCell fromCell, System.Drawing.Point fromOffset, IXlsxCell toCell, System.Drawing.Point toOffset)
    {
        this.Placement = XlsxPicturePlacement.MoveAndSize;

        this.TopLeftCell = fromCell ?? throw new ArgumentNullException(nameof(fromCell));
        this.Markers[XlsxMarkerPosition.TopLeft].Offset = fromOffset;

        this.BottomRightCell = toCell ?? throw new ArgumentNullException(nameof(toCell));
        this.Markers[XlsxMarkerPosition.BottomRight].Offset = toOffset;

        return this;
    }

    public IXlsxPicture Scale(double factor, bool relativeToOriginal = false)
    {
        return this.ScaleHeight(factor, relativeToOriginal).ScaleWidth(factor, relativeToOriginal);
    }

    public IXlsxPicture ScaleHeight(double factor, bool relativeToOriginal = false)
    {
        this.Height = Convert.ToInt32((relativeToOriginal ? this.OriginalHeight : this.Height) * factor);
        return this;
    }

    public IXlsxPicture ScaleWidth(double factor, bool relativeToOriginal = false)
    {
        this.Width = Convert.ToInt32((relativeToOriginal ? this.OriginalWidth : this.Width) * factor);
        return this;
    }

    public IXlsxPicture WithPlacement(XlsxPicturePlacement value)
    {
        this.Placement = value;
        return this;
    }

    public IXlsxPicture WithSize(int width, int height)
    {
        this.Width = width;
        this.Height = height;
        return this;
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        ImageStream.Dispose();
        _disposed = true;
    }

    internal IXlsxPicture CopyTo(XlsxWorksheet targetSheet)
    {
        ThrowIfDisposed();
        targetSheet ??= Worksheet as XlsxWorksheet;

        IXlsxPicture newPicture;
        if (targetSheet == Worksheet)
            newPicture = targetSheet.AddPicture(ImageStream, Format);
        else
            newPicture = targetSheet.AddPicture(ImageStream, Format, Name);

        newPicture = newPicture
                .WithPlacement(XlsxPicturePlacement.FreeFloating)
                .WithSize(Width, Height)
                .WithPlacement(Placement);

        switch (Placement)
        {
            case XlsxPicturePlacement.FreeFloating:
                newPicture.MoveTo(Left, Top);
                break;

            case XlsxPicturePlacement.Move:
                newPicture.MoveTo(targetSheet.Cell(TopLeftCell.Address), GetOffset(XlsxMarkerPosition.TopLeft));
                break;

            case XlsxPicturePlacement.MoveAndSize:
                newPicture.MoveTo(targetSheet.Cell(TopLeftCell.Address), GetOffset(XlsxMarkerPosition.TopLeft), targetSheet.Cell(BottomRightCell.Address),
                    GetOffset(XlsxMarkerPosition.BottomRight));
                break;
        }

        return newPicture;
    }

    internal void SetName(string value)
    {
        if (string.IsNullOrEmpty(value))
            throw new ArgumentException("Picture names cannot be null or empty.");

        _name = value;
    }

    private void Init(XlsxPictureInfo info)
    {
        Format = info.Format;
        Size size = info.GetSizePx(Worksheet.Workbook.DpiX, Worksheet.Workbook.DpiY);
        _width = OriginalWidth = size.Width;
        _height = OriginalHeight = size.Height;
    }

    private void ThrowIfDisposed()
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
    }
}
