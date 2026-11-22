using System.Text;
using System.Xml.Linq;
using System.Xml;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Drawings.Styles;
using Aid.Xlsx.Drawings;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Internal.IO
{
    internal static class VmlDrawingPartWriter
    {
        internal static readonly XNamespace Vml = "urn:schemas-microsoft-com:vml";
        private static readonly XNamespace Office = "urn:schemas-microsoft-com:office:office";
        internal static readonly XNamespace Excel = "urn:schemas-microsoft-com:office:excel";

        internal static XDocument? GenerateContent(Stream? existing, XlsxWorksheet worksheet, bool hasComments)
        {
            var document = (existing is null ? null : XDocumentExtensions.Load(existing)) ?? new XDocument(new XElement("xml"));
            var root = document.Root!.Element("xml") ?? document.Root;
            root.Elements().Where(element => IsCommentShape(element) ||
                (element.Name == Vml + "shapetype" && (string?)element.Attribute("id") == XlsxConstants.Comment.ShapeTypeId)).Remove();
            if (!hasComments)
                return root.HasElements ? document : null;

            var template = new XElement(Vml + "shapetype",
                new XAttribute("id", XlsxConstants.Comment.ShapeTypeId),
                new XAttribute("coordsize", "21600,21600"),
                new XAttribute(Office + "spt", 202),
                new XAttribute("path", "m,l,21600r21600,l21600,xe"),
                new XAttribute(XNamespace.Xmlns + "o", Office),
                new XAttribute(XNamespace.Xmlns + "v", Vml),
                new XElement(Vml + "stroke", new XAttribute("joinstyle", "miter")),
                new XElement(Vml + "path", new XAttribute("gradientshapeok", "true"), new XAttribute(Office + "connecttype", "rect")));

            var cells = worksheet.Internals.CellsCollection.GetCells(cell => cell.HasComment);
            root.AddFirst(template, cells.Select(GenerateCommentShape));
            return document;
        }

        internal static bool IsCommentShape(XElement element)
        {
            if (element.Name != Vml + "shape")
                return false;

            var shapeType = (string?)element.Attribute("type");
            var objectType = (string?)element.Element(Excel + "ClientData")?.Attribute("ObjectType");
            if (objectType is not null)
                return objectType == "Note";

            return shapeType == "#" + XlsxConstants.Comment.ShapeTypeId || shapeType == "#" + XlsxConstants.Comment.AlternateShapeTypeId;
        }

        internal static void Write(Stream target, XDocument document)
        {
            using var writer = XmlWriter.Create(target, new XmlWriterSettings { Encoding = Encoding.UTF8, OmitXmlDeclaration = true });
            document.WriteTo(writer);
        }

        // VML Shape for Comment
        private static XElement GenerateCommentShape(XlsxCell cell)
        {
            var comment = cell.GetComment();
            var style = comment.Style;
            var colors = style.ColorsAndLines;
            var fillColor = "#" + colors.FillColor.Color.ToHex()[2..];
            return new XElement(Vml + "shape",
                new XAttribute("id", "_x0000_s" + comment.ShapeId.ToInvariantString()),
                new XAttribute("style", GetCommentStyle(cell)),
                string.IsNullOrWhiteSpace(style.Web.AlternateText) ? null : new XAttribute("alt", style.Web.AlternateText),
                new XAttribute(Office + "insetmode", style.Margins.Automatic ? "auto" : "custom"),
                new XAttribute("fillcolor", fillColor),
                new XAttribute("strokecolor", "#" + colors.LineColor.Color.ToHex()[2..]),
                new XAttribute("strokeweight", colors.LineWeight.ToInvariantString() + "pt"),
                new XAttribute("type", "#" + XlsxConstants.Comment.ShapeTypeId),
                new XAttribute(XNamespace.Xmlns + "o", Office),
                new XAttribute(XNamespace.Xmlns + "v", Vml),
                new XElement(Vml + "fill", Opacity(colors.FillTransparency), new XAttribute("color2", fillColor)),
                GetStroke(style),
                new XElement(Vml + "shadow", new XAttribute("obscured", "true"), new XAttribute("color", "black")),
                new XElement(Vml + "path", new XAttribute(Office + "connecttype", "none")),
                GetTextBox(style),
                new XElement(Excel + "ClientData", new XAttribute("ObjectType", "Note"), new XAttribute(XNamespace.Xmlns + "xvml", Excel),
                    new XElement(Excel + "MoveWithCells", (style.Properties.Positioning == XlsxDrawingAnchor.Absolute).ToString()),
                    new XElement(Excel + "SizeWithCells", (style.Properties.Positioning != XlsxDrawingAnchor.MoveAndSizeWithCells).ToString()),
                    new XElement(Excel + "Anchor", GetAnchor(cell)),
                    new XElement(Excel + "TextHAlign", style.Alignment.Horizontal.ToString().ToCamel()),
                    new XElement(Excel + "TextVAlign", style.Alignment.Vertical.ToString().ToCamel()),
                    new XElement(Excel + "AutoFill", "False"),
                    new XElement(Excel + "Row", cell.Address.RowNumber - 1),
                    new XElement(Excel + "Column", cell.Address.ColumnNumber - 1),
                    new XElement(Excel + "Locked", style.Protection.Locked.ToString()),
                    new XElement(Excel + "LockText", style.Protection.LockText.ToString()),
                    new XElement(Excel + "Visible", comment.Visible.ToString())));
        }

        private static XElement GetStroke(IXlsxDrawingStyle style)
        {
            var colors = style.ColorsAndLines;
            var dash = colors.LineDash;
            var line = colors.LineStyle switch
            {
                XlsxLineStyle.Single => "single",
                XlsxLineStyle.ThickBetweenThin => "thickBetweenThin",
                XlsxLineStyle.ThickThin => "thickThin",
                XlsxLineStyle.ThinThick => "thinThick",
                XlsxLineStyle.ThinThin => "thinThin",
                _ => throw new ArgumentOutOfRangeException(nameof(style), colors.LineStyle, "Invalid line style.")
            };
            return new XElement(Vml + "stroke", Opacity(colors.LineTransparency), new XAttribute("linestyle", line),
                new XAttribute("dashstyle", dash is XlsxDashStyle.RoundDot or XlsxDashStyle.SquareDot ? "shortDot" : dash.ToString().ToCamel()),
                dash == XlsxDashStyle.RoundDot ? new XAttribute("endcap", "round") : null);
        }

        private static XAttribute? Opacity(double value) => value < 1 ? new XAttribute("opacity", Math.Round(value, 2)) : null;

        private static XElement GetTextBox(IXlsxDrawingStyle ds)
        {
            var sb = new StringBuilder();
            var a = ds.Alignment;

            if (a.Direction == XlsxDrawingTextDirection.Context)
                sb.Append("mso-direction-alt:auto;");
            else if (a.Direction == XlsxDrawingTextDirection.RightToLeft)
                sb.Append("direction:RTL;");

            if (a.Orientation != XlsxDrawingTextOrientation.LeftToRight)
            {
                sb.Append("layout-flow:vertical;");
                if (a.Orientation == XlsxDrawingTextOrientation.BottomToTop)
                    sb.Append("mso-layout-flow-alt:bottom-to-top;");
                else if (a.Orientation == XlsxDrawingTextOrientation.Vertical)
                    sb.Append("mso-layout-flow-alt:top-to-bottom;");
            }
            if (a.AutomaticSize)
                sb.Append("mso-fit-shape-to-text:t;");

            var tb = new XElement(Vml + "textbox");

            if (sb.Length > 0)
                tb.SetAttributeValue("style", sb.ToString());

            var dm = ds.Margins;
            if (!dm.Automatic)
                tb.SetAttributeValue("inset", String.Concat(
                    dm.Left.ToInvariantString(), "in,",
                    dm.Top.ToInvariantString(), "in,",
                    dm.Right.ToInvariantString(), "in,",
                    dm.Bottom.ToInvariantString(), "in"));

            return tb;
        }

        private static string GetAnchor(XlsxCell cell)
        {
            var c = cell.GetComment();
            var cWidth = c.Style.Size.Width;
            var fcNumber = c.Position.Column - 1;
            var fcOffset = Convert.ToInt32(c.Position.ColumnOffset * 7.5);
            var widthFromColumns = cell.Worksheet.Column(c.Position.Column).Width - c.Position.ColumnOffset;
            var lastCell = cell.CellRight(c.Position.Column - cell.Address.ColumnNumber);
            while (widthFromColumns <= cWidth)
            {
                lastCell = lastCell.CellRight();
                widthFromColumns += lastCell.WorksheetColumn().Width;
            }

            var lcNumber = lastCell.WorksheetColumn().ColumnNumber() - 1;
            var lcOffset = Convert.ToInt32((lastCell.WorksheetColumn().Width - (widthFromColumns - cWidth)) * 7.5);

            var cHeight = c.Style.Size.Height; //c.Style.Size.Height * 72.0;
            var frNumber = c.Position.Row - 1;
            var frOffset = Convert.ToInt32(c.Position.RowOffset);
            var heightFromRows = cell.Worksheet.Row(c.Position.Row).Height - c.Position.RowOffset;
            lastCell = cell.CellBelow(c.Position.Row - cell.Address.RowNumber);
            while (heightFromRows <= cHeight)
            {
                lastCell = lastCell.CellBelow();
                heightFromRows += lastCell.WorksheetRow().Height;
            }

            var lrNumber = lastCell.WorksheetRow().RowNumber() - 1;
            var lrOffset = Convert.ToInt32(lastCell.WorksheetRow().Height - (heightFromRows - cHeight));
            return FormattableString.Invariant($"{fcNumber}, {fcOffset}, {frNumber}, {frOffset}, {lcNumber}, {lcOffset}, {lrNumber}, {lrOffset}");
        }

        private static string GetCommentStyle(XlsxCell cell)
        {
            var c = cell.GetComment();
            var sb = new StringBuilder("position:absolute; ");

            sb.Append("visibility:");
            sb.Append(c.Visible ? "visible" : "hidden");
            sb.Append(";");

            sb.Append("width:");
            sb.Append(Math.Round(c.Style.Size.Width * 7.5, 2).ToInvariantString());
            sb.Append("pt;");
            sb.Append("height:");
            sb.Append(Math.Round(c.Style.Size.Height, 2).ToInvariantString());
            sb.Append("pt;");

            sb.Append("z-index:");
            sb.Append(c.ZOrder.ToInvariantString());

            return sb.ToString();
        }

    }
}
