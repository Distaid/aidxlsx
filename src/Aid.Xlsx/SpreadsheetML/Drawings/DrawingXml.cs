using System.Drawing;
using System.Xml;
using System.Xml.Linq;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.SpreadsheetML.Drawings;

internal static class DrawingXml
{
    private static readonly XNamespace SpreadsheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace StrictSpreadsheetDrawing = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing";
    private static readonly XNamespace StrictDrawing = "http://purl.oclc.org/ooxml/drawingml/main";

    internal static XElement Create() => new(SpreadsheetDrawing + "wsDr",
        new XAttribute(XNamespace.Xmlns + "xdr", SpreadsheetDrawing.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "a", Drawing.NamespaceName),
        new XAttribute(XNamespace.Xmlns + "r", OpenXmlConst.RelationshipsNs));

    internal static XElement Read(Stream stream)
    {
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        var root = XElement.Load(reader, System.Xml.Linq.LoadOptions.PreserveWhitespace);
        if (root.Name != SpreadsheetDrawing + "wsDr" && root.Name != StrictSpreadsheetDrawing + "wsDr")
            throw new InvalidDataException("The drawing part must contain a worksheet drawing.");

        return root;
    }

    internal static void Write(Stream stream, XElement root)
    {
        using var writer = XmlWriter.Create(stream);
        writer.WriteStartDocument(standalone: true);
        root.WriteTo(writer);
        writer.WriteEndDocument();
    }

    internal static void LoadPictures(XElement root, XlsxWorksheet sheet, Func<string, Stream?> openImage)
    {
        XNamespace ns = root.Name.Namespace;
        XNamespace a = GetDrawingNamespace(ns);
        XlsxWorkbook workbook = sheet.Workbook;
        foreach (XElement anchor in root.Elements())
        {
            if (anchor.Name.Namespace != ns || !IsPictureAnchor(anchor))
                continue;

            XElement? shape = anchor.Descendants(ns + "pic").FirstOrDefault();
            if (shape is null || ImageId(shape) is not { } id)
                continue;

            XElement? properties = shape.Element(ns + "nvPicPr")?.Element(ns + "cNvPr");
            if ((string?)properties?.Attribute("name") is not { } name || (int?)properties.Attribute("id") is not { } pictureId)
                continue;

            using Stream? stream = openImage(id);
            if (stream is null)
                continue;

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            var picture = (XlsxPicture)sheet.AddPicture(buffer, name, pictureId);
            picture.RelId = id;
            picture.Placement = XlsxPicturePlacement.FreeFloating;

            XElement? transform = shape.Element(ns + "spPr")?.Element(a + "xfrm");
            XElement? size = transform?.Element(a + "ext") ?? anchor.Element(ns + "ext");
            if ((long?)size?.Attribute("cx") is { } width)
                picture.Width = ToPixels(width, workbook.DpiX);
            if ((long?)size?.Attribute("cy") is { } height)
                picture.Height = ToPixels(height, workbook.DpiY);

            switch (anchor.Name.LocalName)
            {
                case "absoluteAnchor":
                    MoveAbsolute(anchor.Element(ns + "pos"));
                    break;
                case "twoCellAnchor" when (string?)anchor.Attribute("editAs") == "absolute":
                    MoveAbsolute(transform?.Element(a + "off"));
                    break;
                default:
                    XlsxMarker from = ReadMarker(sheet, anchor.Element(ns + "from"), ns);
                    if (anchor.Name.LocalName == "twoCellAnchor" && (string?)anchor.Attribute("editAs") != "oneCell")
                    {
                        XlsxMarker to = ReadMarker(sheet, anchor.Element(ns + "to"), ns);
                        picture.MoveTo(from.Cell, from.Offset, to.Cell, to.Offset);
                    }
                    else
                    {
                        picture.MoveTo(from.Cell, from.Offset);
                    }

                    break;
            }

            void MoveAbsolute(XElement? position) => picture.MoveTo(
                ToPixels((long?)position?.Attribute("x") ?? 0, workbook.DpiX),
                ToPixels((long?)position?.Attribute("y") ?? 0, workbook.DpiY));
        }
    }

    internal static string? ImageId(XElement element)
    {
        XNamespace ns = element.Name.Namespace;
        XElement? blip = element.Descendants(ns + "blipFill").FirstOrDefault()?.Element(GetDrawingNamespace(ns) + "blip");
        return (string?)blip?.Attribute(XName.Get("embed", OpenXmlConst.RelationshipsNs))
            ?? (string?)blip?.Attribute(XName.Get("embed", OpenXmlConst.StrictRelationshipsNs));
    }

    internal static bool IsPictureAnchor(XElement element) =>
        element.Name.LocalName is "absoluteAnchor" or "oneCellAnchor" or "twoCellAnchor" &&
        element.Descendants(element.Name.Namespace + "pic").Any();

    internal static XElement CreatePictureAnchor(XNamespace ns, XlsxPicture picture, string imageId)
    {
        XNamespace a = GetDrawingNamespace(ns);
        var r = (XNamespace)(ns == StrictSpreadsheetDrawing ? OpenXmlConst.StrictRelationshipsNs : OpenXmlConst.RelationshipsNs);
        IXlsxWorksheet sheet = picture.Worksheet;
        XlsxWorkbook workbook = sheet.Workbook;
        var width = ToEmu(picture.Width, workbook.DpiX);
        var height = ToEmu(picture.Height, workbook.DpiY);
        var shape = new XElement(ns + "pic",
            new XElement(ns + "nvPicPr",
                new XElement(ns + "cNvPr", new XAttribute("id", 0), new XAttribute("name", picture.Name)),
                new XElement(ns + "cNvPicPr", new XElement(a + "picLocks", new XAttribute("noChangeAspect", "1")))),
            new XElement(ns + "blipFill",
                new XElement(a + "blip", new XAttribute(r + "embed", imageId), new XAttribute("cstate", "print")),
                new XElement(a + "stretch", new XElement(a + "fillRect"))),
            new XElement(ns + "spPr",
                new XElement(a + "xfrm",
                    new XElement(a + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(a + "ext", new XAttribute("cx", width), new XAttribute("cy", height))),
                new XElement(a + "prstGeom", new XAttribute("prst", "rect"))));
        var extent = new XElement(ns + "ext", new XAttribute("cx", width), new XAttribute("cy", height));
        XlsxMarker from = picture.Markers[XlsxMarkerPosition.TopLeft] ?? new XlsxMarker(sheet.Cell("A1"));
        XlsxMarker to = picture.Markers[XlsxMarkerPosition.BottomRight] ?? new XlsxMarker(sheet.Cell("A1"), new Point(picture.Width, picture.Height));

        XElement anchor = picture.Placement switch
        {
            XlsxPicturePlacement.FreeFloating => new XElement(ns + "absoluteAnchor",
                new XElement(ns + "pos", new XAttribute("x", ToEmu(picture.Left, workbook.DpiX)),
                    new XAttribute("y", ToEmu(picture.Top, workbook.DpiY))), extent),
            XlsxPicturePlacement.Move => new XElement(ns + "oneCellAnchor", Marker("from", from), extent),
            XlsxPicturePlacement.MoveAndSize => new XElement(ns + "twoCellAnchor", Marker("from", from), Marker("to", to)),
            _ => throw new ArgumentOutOfRangeException(nameof(picture)),
        };
        anchor.Add(shape, new XElement(ns + "clientData"));
        return anchor;

        XElement Marker(string name, XlsxMarker marker) => new(ns + name,
            new XElement(ns + "col", marker.ColumnNumber - 1),
            new XElement(ns + "colOff", ToEmu(marker.Offset.X, workbook.DpiX)),
            new XElement(ns + "row", marker.RowNumber - 1),
            new XElement(ns + "rowOff", ToEmu(marker.Offset.Y, workbook.DpiY)));
    }

    private static XNamespace GetDrawingNamespace(XNamespace ns) => ns == StrictSpreadsheetDrawing ? StrictDrawing : Drawing;

    private static XlsxMarker ReadMarker(XlsxWorksheet sheet, XElement? marker, XNamespace ns)
    {
        var row = Math.Clamp((long?)marker?.Element(ns + "row") ?? 0, 0, XlsxHelper.MaxRowNumber - 1);
        var column = Math.Clamp((long?)marker?.Element(ns + "col") ?? 0, 0, XlsxHelper.MaxColumnNumber - 1);
        return new XlsxMarker(sheet.Cell((int)row + 1, (int)column + 1), new Point(
            ToPixels((long?)marker?.Element(ns + "colOff") ?? 0, sheet.Workbook.DpiX),
            ToPixels((long?)marker?.Element(ns + "rowOff") ?? 0, sheet.Workbook.DpiY)));
    }

    private static int ToPixels(long emu, double dpi) => Convert.ToInt32(emu * dpi / 914400);

    private static long ToEmu(int pixels, double dpi) => Convert.ToInt64(914400L * pixels / dpi);
}
