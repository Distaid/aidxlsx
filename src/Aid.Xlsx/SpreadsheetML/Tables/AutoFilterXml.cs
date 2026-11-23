using System.Globalization;
using System.Xml.Linq;
using Aid.Xlsx.Model.AutoFilters;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Grid.Ranges.Sort;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.SpreadsheetML.Tables;

internal static class AutoFilterXml
{
    private static readonly XNamespace Spreadsheet = OpenXmlConst.Main2006SsNs;
    private static readonly string[] DateParts = ["year", "month", "day", "hour", "minute", "second"];

    internal static void LoadWorksheetFilter(XElement source, XlsxWorksheet worksheet)
    {
        GetRange(worksheet, source).SetAutoFilter();
        XlsxAutoFilter filter = worksheet.AutoFilter;
        ReadColumns(source, filter);

        XNamespace ns = source.Name.Namespace;
        if (source.Element(ns + "sortState")?.Element(ns + "sortCondition") is { } condition)
        {
            var sortedColumn = GetRange(worksheet, condition).FirstCell().Address.ColumnNumber;
            filter.SortColumn = sortedColumn - filter.Range!.FirstCell().Address.ColumnNumber + 1;
            filter.Sorted = true;
            filter.SortOrder = (bool?)condition.Attribute("descending") == true ? XlsxSortOrder.Descending : XlsxSortOrder.Ascending;
        }
    }

    internal static void ReadColumns(XElement source, XlsxAutoFilter filter)
    {
        XNamespace ns = source.Name.Namespace;
        foreach (XElement element in source.Elements(ns + "filterColumn"))
        {
            XlsxFilterColumn column = filter.Column(checked((int)(uint)element.Attribute("colId")! + 1));
            if (element.Element(ns + "customFilters") is { } custom)
                ReadCustom(custom, column);
            else if (element.Element(ns + "filters") is { } regular)
                ReadRegular(regular, column);
            else if (element.Element(ns + "top10") is { } top)
                ReadTopBottom(top, column);
            else if (element.Element(ns + "dynamicFilter") is { } dynamic)
                ReadDynamic(dynamic, column);
        }
    }

    private static void ReadCustom(XElement source, XlsxFilterColumn column)
    {
        column.FilterType = XlsxFilterType.Custom;
        XlsxConnector connector = (bool?)source.Attribute("and") == true ? XlsxConnector.And : XlsxConnector.Or;
        foreach (XElement element in source.Elements(source.Name.Namespace + "customFilter"))
        {
            XlsxFilterOperator op = (string?)element.Attribute("operator") switch
            {
                null or "equal" => XlsxFilterOperator.Equal,
                "notEqual" => XlsxFilterOperator.NotEqual,
                "greaterThan" => XlsxFilterOperator.GreaterThan,
                "greaterThanOrEqual" => XlsxFilterOperator.EqualOrGreaterThan,
                "lessThan" => XlsxFilterOperator.LessThan,
                "lessThanOrEqual" => XlsxFilterOperator.EqualOrLessThan,
                var value => throw new InvalidDataException($"Unknown filter operator '{value}'.")
            };
            var operand = Required(element, "val");
            XlsxFilter filter = op is XlsxFilterOperator.Equal or XlsxFilterOperator.NotEqual
                ? XlsxFilter.CreateCustomPatternFilter(operand, op == XlsxFilterOperator.Equal, connector)
                : XlsxFilter.CreateCustomFilter(XlsxCellValue.FromText(operand, CultureInfo.InvariantCulture), op, connector);
            column.AddFilter(filter);
        }
    }

    private static void ReadRegular(XElement source, XlsxFilterColumn column)
    {
        column.FilterType = XlsxFilterType.Regular;
        XNamespace ns = source.Name.Namespace;
        foreach (XElement element in source.Elements(ns + "filter"))
            column.AddFilter(XlsxFilter.CreateRegularFilter(Required(element, "val")));

        Span<int> date = stackalloc int[6];
        foreach (XElement element in source.Elements(ns + "dateGroupItem"))
        {
            var grouping = Array.IndexOf(DateParts, (string?)element.Attribute("dateTimeGrouping"));
            if (grouping < 0)
                continue;

            date.Clear();
            date[1] = 1;
            date[2] = 1;
            var i = 0;
            for (; i <= grouping && (int?)element.Attribute(DateParts[i]) is { } value; i++)
                date[i] = value;

            if (i > grouping)
                column.AddFilter(XlsxFilter.CreateDateGroupFilter(new DateTime(date[0], date[1], date[2], date[3], date[4], date[5]), (XlsxDateTimeGrouping)grouping));
        }
    }

    private static void ReadTopBottom(XElement source, XlsxFilterColumn column)
    {
        var takeTop = (bool?)source.Attribute("top") ?? true;
        var value = (int)(double)source.Attribute("val")!;
        column.FilterType = XlsxFilterType.TopBottom;
        column.TopBottomType = (bool?)source.Attribute("percent") == true ? XlsxTopBottomType.Percent : XlsxTopBottomType.Items;
        column.TopBottomPart = takeTop ? XlsxTopBottomPart.Top : XlsxTopBottomPart.Bottom;
        column.TopBottomValue = value;
        column.AddFilter(XlsxFilter.CreateTopBottom(takeTop, value));
    }

    private static void ReadDynamic(XElement source, XlsxFilterColumn column)
    {
        XlsxFilterDynamicType type = (string?)source.Attribute("type") switch
        {
            null or "aboveAverage" => XlsxFilterDynamicType.AboveAverage,
            "belowAverage" => XlsxFilterDynamicType.BelowAverage,
            var value => throw new InvalidDataException($"Unsupported dynamic filter type '{value}'.")
        };
        var threshold = (double)source.Attribute("val")!;
        column.FilterType = XlsxFilterType.Dynamic;
        column.DynamicType = type;
        column.DynamicValue = threshold;
        column.AddFilter(XlsxFilter.CreateAverage(threshold, type == XlsxFilterDynamicType.AboveAverage));
    }

    internal static XElement Write(XlsxAutoFilter filter)
    {
        IXlsxRange range = filter.Range!;
        var target = new XElement(Spreadsheet + "autoFilter", new XAttribute("ref", range.RangeAddress.ToString()!));
        foreach ((int index, XlsxFilterColumn? column) in filter.Columns)
        {
            XElement definition = column.FilterType switch
            {
                XlsxFilterType.Custom => WriteCustom(column),
                XlsxFilterType.TopBottom => new XElement(Spreadsheet + "top10",
                    column.TopBottomPart == XlsxTopBottomPart.Bottom ? new XAttribute("top", "0") : null,
                    column.TopBottomType == XlsxTopBottomType.Percent ? new XAttribute("percent", "1") : null,
                    new XAttribute("val", column.TopBottomValue)),
                XlsxFilterType.Dynamic => new XElement(Spreadsheet + "dynamicFilter",
                    new XAttribute("type", column.DynamicType == XlsxFilterDynamicType.AboveAverage ? "aboveAverage" : "belowAverage"),
                    new XAttribute("val", column.DynamicValue)),
                XlsxFilterType.Regular => WriteRegular(column),
                _ => throw new NotSupportedException()
            };
            target.Add(new XElement(Spreadsheet + "filterColumn", new XAttribute("colId", index - 1), definition));
        }

        if (filter.Sorted)
        {
            IXlsxRange sortRange = range.RowCount() > 1 ? range.Range(range.FirstCell().CellBelow(), range.LastCell()) : range;
            IXlsxRange columnRange = range.Range(1, filter.SortColumn, range.RowCount(), filter.SortColumn);
            target.Add(new XElement(Spreadsheet + "sortState", new XAttribute("ref", sortRange.RangeAddress.ToString()!),
                new XElement(Spreadsheet + "sortCondition",
                    filter.SortOrder == XlsxSortOrder.Descending ? new XAttribute("descending", "1") : null,
                    new XAttribute("ref", columnRange.RangeAddress.ToString()!))));
        }

        return target;
    }

    private static XElement WriteCustom(XlsxFilterColumn column)
    {
        var element = new XElement(Spreadsheet + "customFilters");
        foreach (XlsxFilter filter in column)
        {
            if (filter.Connector == XlsxConnector.And)
                element.SetAttributeValue("and", "1");

            var op = filter.Operator switch
            {
                XlsxFilterOperator.Equal => null,
                XlsxFilterOperator.NotEqual => "notEqual",
                XlsxFilterOperator.GreaterThan => "greaterThan",
                XlsxFilterOperator.EqualOrGreaterThan => "greaterThanOrEqual",
                XlsxFilterOperator.LessThan => "lessThan",
                XlsxFilterOperator.EqualOrLessThan => "lessThanOrEqual",
                _ => throw new NotSupportedException()
            };
            element.Add(new XElement(Spreadsheet + "customFilter", op is null ? null : new XAttribute("operator", op),
                new XAttribute("val", filter.CustomValue.ToString(CultureInfo.InvariantCulture))));
        }

        return element;
    }

    private static XElement WriteRegular(XlsxFilterColumn column)
    {
        var element = new XElement(Spreadsheet + "filters");
        foreach (XlsxFilter filter in column)
        {
            if (filter.Value is string text)
                element.Add(new XElement(Spreadsheet + "filter", new XAttribute("val", text)));
        }

        Span<int> date = stackalloc int[6];
        foreach (XlsxFilter filter in column)
        {
            if (filter.Value is not DateTime value)
                continue;

            date[0] = value.Year;
            date[1] = value.Month;
            date[2] = value.Day;
            date[3] = value.Hour;
            date[4] = value.Minute;
            date[5] = value.Second;
            var grouping = (int)filter.DateTimeGrouping;
            var item = new XElement(Spreadsheet + "dateGroupItem");
            for (var i = 0; i <= grouping; i++)
                item.SetAttributeValue(DateParts[i], date[i]);
            item.SetAttributeValue("dateTimeGrouping", DateParts[grouping]);
            element.Add(item);
        }

        return element;
    }

    private static string Required(XElement element, string attribute) => (string?)element.Attribute(attribute)
        ?? throw new InvalidDataException($"Missing '{attribute}' on '{element.Name.LocalName}'.");

    private static IXlsxRange GetRange(XlsxWorksheet worksheet, XElement element)
    {
        var reference = Required(element, "ref");
        return worksheet.Range(reference) ?? throw new InvalidDataException($"Invalid filter range '{reference}'.");
    }
}
