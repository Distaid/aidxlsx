# Images and drawings

A picture is a drawing anchored to a worksheet. The API can add pictures, place them
relative to cells or to the sheet origin, resize them, rename, copy, and delete them.
It does not create shapes or charts; those are preserved as package data when they
already exist. This guide covers pictures, anchors, media storage, and the drawing
features that are intentionally read-only.

## Pictures

Use `AddPicture` when an image should appear on a sheet, for example a logo, a product
photo, or a rendered chart exported as PNG. A picture is anchored immediately; the
default placement moves and resizes with its cells.

### Add a picture from a file

```csharp
using System;
using System.IO;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Branding");

IXlsxPicture logo = sheet.AddPicture("logo.png", "Logo");
logo.MoveTo(sheet.Cell("B2"), 8, 6);
logo.WithSize(240, 80);

workbook.SaveAs("branding.xlsx");
```

The file is read once and copied into memory, so the source file can be changed or
deleted after `AddPicture` returns. The second argument is the picture name; without
it the sheet assigns `Picture 1`, `Picture 2`, and so on.

### Add a picture from a stream

```csharp
using System;
using System.IO;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Scan");

using FileStream stream = File.OpenRead("scan.jpg");
IXlsxPicture scan = sheet.AddPicture(stream, XlsxPictureFormat.Jpeg, "Scan");
scan.MoveTo(sheet.Cell("A1"));
```

The stream must be seekable: the reader rewinds it to inspect the header and copies
from position zero. The optional `XlsxPictureFormat` is passed to the graphic engine
as a hint; the default engine detects the actual format from the bytes. Add the
picture before disposing the stream.

### Supported formats

| `XlsxPictureFormat` | Detected by default engine | Media part extension |
| --- | --- | --- |
| `Png` | Yes | `.png` |
| `Jpeg` | Yes | `.jpeg` |
| `Gif` | Yes | `.gif` |
| `Bmp` | Yes | `.bmp` |
| `Tiff` | Yes | `.tiff` |
| `Emf` | Yes | `.emf` |
| `Wmf` | Yes | `.wmf` |
| `Webp` | Yes | `.webp` |
| `Pcx` | Yes | `.pcx` |
| `Icon` | No reader in the default engine | `.ico` if supplied by a custom engine |
| `Unknown` | Not a valid target | Saving such a picture throws |

`AddPicture` throws `ArgumentException` when no reader recognises the stream. If you
implement `IXlsxGraphicEngine`, your `GetPictureInfo` implementation decides which
formats are accepted and may honour the format argument.

### Placement, size, and scaling

`XlsxPicturePlacement` controls the anchor type and the resize behaviour.

| Placement | Anchor | Resizes with cells |
| --- | --- | --- |
| `MoveAndSize` (default) | Two-cell (`twoCellAnchor`) | Yes |
| `Move` | One-cell (`oneCellAnchor`) | No, keeps its pixel size |
| `FreeFloating` | Absolute (`absoluteAnchor`) | No, ignores the grid |

```csharp
// Two-cell anchor: the picture spans B2 to H12 and resizes with those cells.
var banner = sheet.AddPicture("banner.jpg", "Banner");
banner.MoveTo(sheet.Cell("B2"), sheet.Cell("H12"));

// One-cell anchor: top-left at C3 plus 4 x 4 px, fixed pixel size.
var badge = sheet.AddPicture("badge.png", "Badge");
badge.MoveTo(sheet.Cell("C3"), 4, 4);
badge.WithSize(96, 96);

// Absolute anchor: 600 px from the left, 40 px from the top of the sheet.
var stamp = sheet.AddPicture("stamp.png", "Stamp");
stamp.MoveTo(600, 40);
stamp.WithSize(120, 120);

stamp.Scale(0.5);                            // 120 x 120 -> 60 x 60
stamp.ScaleWidth(1.5, relativeToOriginal: true);
```

`Width` and `Height` are in pixels. Setting them, calling `WithSize`, or calling
`Scale` throws `ArgumentException` while the picture uses `MoveAndSize`, because a
two-cell anchor derives the size from the two cells. Switch to `Move` or
`FreeFloating` first. `OriginalWidth` and `OriginalHeight` keep the decoded image size
and are not changed by scaling; `Scale(..., relativeToOriginal: true)` calculates from
them.

`Left` and `Top` are pixel offsets from the sheet origin and can only be set on a
`FreeFloating` picture, which is what `MoveTo(left, top)` sets up. Offsets inside the
anchor cell are read with `GetOffset`:

```csharp
System.Drawing.Point offset = logo.GetOffset(XlsxMarkerPosition.TopLeft);
Console.WriteLine($"{logo.TopLeftCell.Address.ToStringRelative()} + {offset}");
```

### Anchors and cell edits

`MoveTo(fromCell, toCell)` sets `MoveAndSize`; `MoveTo(cell[, offset])` sets `Move`;
`MoveTo(left, top)` sets `FreeFloating`. `WithPlacement` changes only the placement
flag, so use a `MoveTo` overload when you also need anchors. Anchor cells must belong
to the same worksheet, otherwise the call throws `InvalidOperationException`.
`IXlsxPicture` does not extend `IXlsxDrawing<T>`, so a picture has no `Position`,
`ZOrder`, or `Style` member; placement and offsets are the only positioning API.

Anchors are tracked as cells plus pixel offsets, so inserting or deleting rows and
columns moves a picture with its anchor. Resizing a row or column changes a
`MoveAndSize` picture but not a `Move` or `FreeFloating` one. A picture's size is
stored in EMU: 914,400 EMU per inch. At the default 96 DPI, 96 px is exactly one inch.
The conversion uses `LoadOptions.Dpi`, so the same pixel size maps to a different EMU
size when the workbook is loaded with a different DPI. At DPI values that do not
divide 914,400 evenly, a load and save can change a dimension by one pixel.

### Names, listing, and deletion

```csharp
using System;
using Aid.Xlsx.Model.Drawings;

Console.WriteLine(sheet.Pictures.Count);
foreach (IXlsxPicture picture in sheet.Pictures)
    Console.WriteLine($"{picture.Name}: {picture.Width}x{picture.Height}");

if (sheet.Pictures.Contains("Logo"))
{
    IXlsxPicture existing = sheet.Picture("Logo");
    existing.Name = "Company logo";
}

IXlsxPicture duplicate = sheet.Picture("Company logo").Duplicate();
duplicate.MoveTo(10, 10);

sheet.Pictures.Delete("Company logo");
```

`Pictures` is the per-worksheet collection. `Picture` and `Delete` throw
`ArgumentOutOfRangeException` when the name is absent; `TryGetPicture` and `Contains`
are the non-throwing checks. Names are unique within one worksheet, compared without
case, and cannot be null or empty. `Duplicate` copies within the sheet and picks a
fresh automatic name; `CopyTo` copies to another sheet and keeps the name, so a name
collision on the target sheet throws. `picture.Delete()` is equivalent to deleting by
name.

### Z-order and cropping

There is no public z-order property and no crop API. Pictures are written in the order
they appear in the collection, so a newly added picture is drawn above pictures added
before it. To control stacking, add pictures in the intended order. Cropping must be
done in the image file before it is added; the original bytes are stored unchanged.

### Media parts, memory, and preservation

- Each picture keeps its encoded bytes in memory in `ImageStream`. Loading and adding
  both copy the full stream; nothing is decoded into a bitmap and nothing is
  re-encoded on save.
- Every picture is written to a media part under `/xl/media/` with the extension and
  content type of its format. The bytes are the original file bytes.
- Newly added pictures are never deduplicated: adding the same file twice creates two
  media parts. Pictures that shared one media part in a loaded file keep sharing it.
  Deleting a picture removes its anchor and unlinks the media part only when no other
  anchor still references it.
- The cost of a workbook in memory is roughly the sum of the encoded image sizes. This
  matters when a report embeds many high-resolution photos; resize or compress them
  before adding.
- When a worksheet has no pictures and no picture was deleted in this session, its
  drawing part is left untouched. Otherwise the drawing XML is rewritten: picture
  anchors are replaced, unsupported anchors such as charts, shapes, and connectors are
  preserved as raw elements, and drawing IDs are renumbered. The part is not
  guaranteed to stay byte-identical.
- `SaveOptions.ValidatePackage` checks relationships and the structure of supported
  XML parts; it does not validate image content or drawing semantics.

### Rendering and image metadata

`DefaultGraphicEngine.GetPictureInfo` reads a picture's format and dimensions without
adding it to a worksheet. This is useful for pre-flight checks of uploaded files and
for calculating layout before insertion.

```csharp
using System;
using System.IO;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Rendering;
using Aid.Xlsx.Rendering.Images;

using FileStream stream = File.OpenRead("logo.png");
XlsxPictureInfo info = DefaultGraphicEngine.Instance.Value
    .GetPictureInfo(stream, XlsxPictureFormat.Unknown);
Console.WriteLine($"{info.Format}: {info.SizePx.Width}x{info.SizePx.Height} px");
```

`XlsxPictureInfo` reports `Format`, `SizePx`, `SizePhys` (hundredths of a millimetre,
used by vector formats), and `DpiX`/`DpiY`. `DpiX` and `DpiY` are the resolution
embedded in the image, and zero means the image does not carry one. When `SizePx` is
empty, the picture size is computed from `SizePhys` and the workbook DPI.
`GetPictureInfo` throws `ArgumentException` for unrecognised content.

The engine also measures text, so the same `LoadOptions.GraphicEngine` affects image
dimensions and cell layout. See [NativeAOT and trimming](native-aot.md) for supplying
fonts without system access, and [Performance](performance.md) for the DPI option.

## Shapes and charts

### Shapes

Aid.Xlsx has no public API to add a text box, rectangle, line, arrow, or other shape.
The drawing style interfaces (`IXlsxDrawingStyle` and its parts) describe the shape
behind a comment and are not a shape factory. Existing shapes remain in the drawing
part and are preserved as raw XML, but they cannot be created, listed, or edited
through the object model.

### Charts

`IXlsxChart`, `XlsxChartType`, and `IXlsxCharts` are public types, but there is no
public factory for a chart and no public worksheet property that returns an
`IXlsxCharts` instance. A chart can therefore not be created or read through the API.
Charts already present in a workbook are preserved as package data, including when
other pictures on the same drawing part are edited. Deleting or replacing a chart is a
package-level operation outside this API. See [Compatibility](compatibility.md) for the
feature matrix.

### Unsupported drawing parts

Drawing parts can contain more than pictures, charts, and shapes: connectors, group
shapes, text effects, and custom XML extensions are common. The reader keeps the
complete drawing XML and rewrites only the picture anchors it manages. Elements it
does not model survive that rewrite, but markup can be re-serialised, namespace
prefixes can change, and `cNvPr` IDs are renumbered when pictures are present. Do not
expect an unchanged drawing part unless the worksheet's pictures are untouched.

## Limits

| Item | Limit |
| --- | --- |
| Picture name | Unique per worksheet, case-insensitive, non-empty |
| Picture dimensions | Whole pixels; values above `int.MaxValue` are rejected |
| Formats | Detected formats listed above; others throw at add time |
| Cropping | Not exposed |
| Z-order | Not exposed; collection order controls stacking |
| Shape creation | Not exposed |
| Chart creation and reading | Not exposed, preserved only |

Pictures are copied with a worksheet when the sheet is copied, and `Duplicate` copies
one picture within its sheet. Every copy writes the encoded bytes into a new media
part; image bytes are not shared between pictures created through the API. Add a
picture once per sheet that needs it, and prefer compressing large images before
insertion.
