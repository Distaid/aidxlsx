using System.Xml.Linq;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Columns;
using Aid.Xlsx.Model.PageSetup;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.SpreadsheetML.Worksheets;

internal sealed partial class WorksheetXml
{
    internal void UpdateColumns(XlsxWorksheet sheet, XlsxWorkbook.SaveContext context)
    {
        var worksheetStyle = context.GetStyleId(sheet.FormatValue);
        XlsxColumnsCollection model = sheet.Internals.ColumnsCollection;
        if (sheet.Internals.CellsCollection.IsEmpty && model.Count == 0 && worksheetStyle == 0)
        {
            Replace("cols", []);
            return;
        }

        XElement columns = GetOrAdd("cols");
        var existing = columns.Elements(Main + "col").ToDictionary(column => (int)column.Attribute("min")!);
        var lastModelColumn = model.Count == 0 ? 0 : model.Keys.Max();
        var width = ColumnWidth(sheet.ColumnWidth);
        for (var index = 1; index <= lastModelColumn; index++)
        {
            XElement? old = existing.GetValueOrDefault(index);
            XElement column = old is null ? new XElement(Main + "col") : new XElement(old);
            if (old is not null)
            {
                if ((int)old.Attribute("max")! > index)
                {
                    old.SetAttributeValue("min", index + 1);
                    existing.Add(index + 1, old);
                }
                else
                {
                    old.Remove();
                }
            }

            model.TryGetValue(index, out XlsxColumn? value);
            column.SetAttributeValue("min", index);
            column.SetAttributeValue("max", index);
            column.SetAttributeValue("style", value?.FormatValue is { } format ? context.GetStyleId(format) : worksheetStyle);
            column.SetAttributeValue("width", value is null ? width : ColumnWidth(value.Width));
            Flag(column, "customWidth", true);
            Flag(column, "hidden", value?.IsHidden == true, false);
            Flag(column, "collapsed", value?.Collapsed == true, false);
            column.SetAttributeValue("outlineLevel", value?.OutlineLevel is > 0 and var level ? level : (int?)null);
            columns.Add(column);
            existing[index] = column;
        }

        var lastColumn = lastModelColumn;
        foreach (XElement? column in existing.Where(pair => pair.Key > lastModelColumn).Select(pair => pair.Value))
        {
            column.SetAttributeValue("style", worksheetStyle);
            column.SetAttributeValue("width", width);
            Flag(column, "customWidth", true);
            lastColumn = Math.Max(lastColumn, (int)column.Attribute("max")!);
        }

        if (lastColumn < XlsxHelper.MaxColumnNumber && worksheetStyle != 0)
        {
            columns.Add(new XElement(Main + "col", new XAttribute("min", lastColumn + 1), new XAttribute("max", XlsxHelper.MaxColumnNumber),
                new XAttribute("style", worksheetStyle), new XAttribute("width", width), new XAttribute("customWidth", "1")));
        }

        // Compress adjacent equal column definitions while keeping additional attributes from the original part.
        XElement[] ordered = [.. columns.Elements(Main + "col").OrderBy(column => (int)column.Attribute("min")!)];
        ordered.Remove();
        XElement? previous = null;
        foreach (XElement? column in ordered)
        {
            if (previous is not null && (int)previous.Attribute("max")! + 1 == (int)column.Attribute("min")! && ColumnsEqual(previous, column))
            {
                previous.SetAttributeValue("max", (int)column.Attribute("max")!);
            }
            else
            {
                columns.Add(column);
                previous = column;
            }
        }

        if (!columns.HasElements)
            columns.Remove();
    }

    private static bool ColumnsEqual(XElement left, XElement right)
    {
        var leftWidth = (double?)left.Attribute("width");
        var rightWidth = (double?)right.Attribute("width");
        if (leftWidth != rightWidth && !(Math.Abs((leftWidth - rightWidth) ?? double.PositiveInfinity) < XlsxHelper.Epsilon))
            return false;

        static bool IsColumnSetting(XAttribute attribute) => attribute.Name.LocalName is not ("min" or "max" or "width");
        return left.Attributes().Count(IsColumnSetting) == right.Attributes().Count(IsColumnSetting)
            && left.Attributes().Where(IsColumnSetting).All(attribute => (string?)right.Attribute(attribute.Name) == attribute.Value);
    }

    private static double ColumnWidth(double width) => Math.Clamp(width + XlsxConstants.ColumnWidthOffset, 0, 255).SaveRound();

    internal void UpdatePrintSettings(XlsxWorksheet sheet)
    {
        IXlsxPageSetup setup = sheet.PageSetup;
        XElement printOptions = GetOrAdd("printOptions");
        Flag(printOptions, "horizontalCentered", setup.CenterHorizontally);
        Flag(printOptions, "verticalCentered", setup.CenterVertically);
        Flag(printOptions, "headings", setup.ShowRowAndColumnHeadings);
        Flag(printOptions, "gridLines", setup.ShowGridlines);

        XElement margins = GetOrAdd("pageMargins");
        margins.SetAttributeValue("left", setup.Margins.Left);
        margins.SetAttributeValue("right", setup.Margins.Right);
        margins.SetAttributeValue("top", setup.Margins.Top);
        margins.SetAttributeValue("bottom", setup.Margins.Bottom);
        margins.SetAttributeValue("header", setup.Margins.Header);
        margins.SetAttributeValue("footer", setup.Margins.Footer);

        XElement page = GetOrAdd("pageSetup");
        page.SetAttributeValue("orientation", setup.PageOrientation switch
        {
            XlsxPageOrientation.Default => "default",
            XlsxPageOrientation.Portrait => "portrait",
            XlsxPageOrientation.Landscape => "landscape",
            _ => throw new ArgumentOutOfRangeException(nameof(sheet)),
        });
        page.SetAttributeValue("paperSize", (uint)setup.PaperSize);
        Flag(page, "blackAndWhite", setup.BlackAndWhite);
        Flag(page, "draft", setup.DraftQuality);
        page.SetAttributeValue("pageOrder", setup.PageOrder switch
        {
            XlsxPageOrderValues.DownThenOver => "downThenOver",
            XlsxPageOrderValues.OverThenDown => "overThenDown",
            _ => throw new ArgumentOutOfRangeException(nameof(sheet)),
        });
        page.SetAttributeValue("cellComments", setup.ShowComments switch
        {
            XlsxShowCommentsValues.None => "none",
            XlsxShowCommentsValues.AsDisplayed => "asDisplayed",
            XlsxShowCommentsValues.AtEnd => "atEnd",
            _ => throw new ArgumentOutOfRangeException(nameof(sheet)),
        });
        page.SetAttributeValue("errors", setup.PrintErrorValue switch
        {
            XlsxPrintErrorValues.Blank => "blank",
            XlsxPrintErrorValues.Dash => "dash",
            XlsxPrintErrorValues.Displayed => "displayed",
            XlsxPrintErrorValues.NA => "NA",
            _ => throw new ArgumentOutOfRangeException(nameof(sheet)),
        });

        // Negative page numbers are encoded as unsigned integers in the file.
        page.SetAttributeValue("firstPageNumber", setup.FirstPageNumber is { } first ? unchecked((uint)first) : (uint?)null);
        Flag(page, "useFirstPageNumber", setup.FirstPageNumber.HasValue, false);
        page.SetAttributeValue("horizontalDpi", Positive(setup.HorizontalDpi));
        page.SetAttributeValue("verticalDpi", Positive(setup.VerticalDpi));
        page.SetAttributeValue("scale", Positive(setup.Scale));
        if (setup.Scale > 0)
        {
            page.SetAttributeValue("fitToWidth", null);
            page.SetAttributeValue("fitToHeight", null);
        }
        else
        {
            if (setup.PagesWide >= 0 && setup.PagesWide != 1)
                page.SetAttributeValue("fitToWidth", setup.PagesWide);
            if (setup.PagesTall >= 0 && setup.PagesTall != 1)
                page.SetAttributeValue("fitToHeight", setup.PagesTall);
        }
        if ((uint?)page.Attribute("copies") is 0)
            page.SetAttributeValue("copies", null);

        XElement headerFooter = GetOrAdd("headerFooter");
        if (((XlsxHeaderFooter)setup.Header).Changed || ((XlsxHeaderFooter)setup.Footer).Changed)
        {
            headerFooter.RemoveNodes();
            Flag(headerFooter, "scaleWithDoc", setup.ScaleHFWithDocument);
            Flag(headerFooter, "alignWithMargins", setup.AlignHFWithMargins);
            Flag(headerFooter, "differentFirst", setup.DifferentFirstPageOnHF);
            Flag(headerFooter, "differentOddEven", setup.DifferentOddEvenPagesOnHF);
            WriteHeaderFooter("odd", XlsxHFOccurrence.OddPages);
            WriteHeaderFooter("even", XlsxHFOccurrence.EvenPages);
            WriteHeaderFooter("first", XlsxHFOccurrence.FirstPage);
        }

        UpdateBreaks("rowBreaks", setup.RowBreaks, sheet.RangeAddress.LastAddress.RowNumber);
        UpdateBreaks("colBreaks", setup.ColumnBreaks, sheet.RangeAddress.LastAddress.ColumnNumber);

        static int? Positive(int value) => value > 0 ? value : null;

        void WriteHeaderFooter(string prefix, XlsxHFOccurrence occurrence)
        {
            headerFooter.Add(new XElement(Main + (prefix + "Header"), setup.Header.GetText(occurrence)));
            headerFooter.Add(new XElement(Main + (prefix + "Footer"), setup.Footer.GetText(occurrence)));
        }
    }

    private void UpdateBreaks(string name, ICollection<int> breaks, int last)
    {
        if (breaks.Count == 0)
        {
            Replace(name, []);
            return;
        }

        XElement element = GetOrAdd(name);
        var ids = new HashSet<int>();
        foreach (XElement? existing in element.Elements(Main + "brk").ToArray())
        {
            if ((int?)existing.Attribute("id") is { } id && breaks.Contains(id))
                ids.Add(id);
            else
                existing.Remove();
        }

        element.SetAttributeValue("count", breaks.Count);
        element.SetAttributeValue("manualBreakCount", breaks.Count);
        foreach (var number in breaks.Where(number => !ids.Contains(number)))
            element.Add(new XElement(Main + "brk", new XAttribute("id", number), new XAttribute("max", last), new XAttribute("man", "1")));
    }
}
