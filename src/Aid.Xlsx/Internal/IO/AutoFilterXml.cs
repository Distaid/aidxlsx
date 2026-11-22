using System;
using System.Globalization;
using System.IO;
using System.Xml.Linq;

namespace Aid.Xlsx.Internal.IO;

internal static class AutoFilterXml
{
    private static readonly XNamespace Spreadsheet = OpenXmlConst.Main2006SsNs;
    private static readonly string[] DateParts = ["year", "month", "day", "hour", "minute", "second"];

    internal static void LoadWorksheetFilter(XElement source, XLWorksheet worksheet)
    {
        GetRange(worksheet, source).SetAutoFilter();
        var filter = worksheet.AutoFilter;
        ReadColumns(source, filter);

        var ns = source.Name.Namespace;
        if (source.Element(ns + "sortState")?.Element(ns + "sortCondition") is { } condition)
        {
            var sortedColumn = GetRange(worksheet, condition).FirstCell().Address.ColumnNumber;
            filter.SortColumn = sortedColumn - filter.Range!.FirstCell().Address.ColumnNumber + 1;
            filter.Sorted = true;
            filter.SortOrder = (bool?)condition.Attribute("descending") == true ? XLSortOrder.Descending : XLSortOrder.Ascending;
        }
    }

    internal static void ReadColumns(XElement source, XLAutoFilter filter)
    {
        var ns = source.Name.Namespace;
        foreach (var element in source.Elements(ns + "filterColumn"))
        {
            var column = filter.Column(checked((int)(uint)element.Attribute("colId")! + 1));
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

    private static void ReadCustom(XElement source, XLFilterColumn column)
    {
        column.FilterType = XLFilterType.Custom;
        var connector = (bool?)source.Attribute("and") == true ? XLConnector.And : XLConnector.Or;
        foreach (var element in source.Elements(source.Name.Namespace + "customFilter"))
        {
            var op = (string?)element.Attribute("operator") switch
            {
                null or "equal" => XLFilterOperator.Equal,
                "notEqual" => XLFilterOperator.NotEqual,
                "greaterThan" => XLFilterOperator.GreaterThan,
                "greaterThanOrEqual" => XLFilterOperator.EqualOrGreaterThan,
                "lessThan" => XLFilterOperator.LessThan,
                "lessThanOrEqual" => XLFilterOperator.EqualOrLessThan,
                var value => throw new InvalidDataException($"Unknown filter operator '{value}'.")
            };
            var operand = Required(element, "val");
            var filter = op is XLFilterOperator.Equal or XLFilterOperator.NotEqual
                ? XLFilter.CreateCustomPatternFilter(operand, op == XLFilterOperator.Equal, connector)
                : XLFilter.CreateCustomFilter(XLCellValue.FromText(operand, CultureInfo.InvariantCulture), op, connector);
            column.AddFilter(filter);
        }
    }

    private static void ReadRegular(XElement source, XLFilterColumn column)
    {
        column.FilterType = XLFilterType.Regular;
        var ns = source.Name.Namespace;
        foreach (var element in source.Elements(ns + "filter"))
            column.AddFilter(XLFilter.CreateRegularFilter(Required(element, "val")));

        Span<int> date = stackalloc int[6];
        foreach (var element in source.Elements(ns + "dateGroupItem"))
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
                column.AddFilter(XLFilter.CreateDateGroupFilter(new DateTime(date[0], date[1], date[2], date[3], date[4], date[5]), (XLDateTimeGrouping)grouping));
        }
    }

    private static void ReadTopBottom(XElement source, XLFilterColumn column)
    {
        var takeTop = (bool?)source.Attribute("top") ?? true;
        var value = (int)(double)source.Attribute("val")!;
        column.FilterType = XLFilterType.TopBottom;
        column.TopBottomType = (bool?)source.Attribute("percent") == true ? XLTopBottomType.Percent : XLTopBottomType.Items;
        column.TopBottomPart = takeTop ? XLTopBottomPart.Top : XLTopBottomPart.Bottom;
        column.TopBottomValue = value;
        column.AddFilter(XLFilter.CreateTopBottom(takeTop, value));
    }

    private static void ReadDynamic(XElement source, XLFilterColumn column)
    {
        var type = (string?)source.Attribute("type") switch
        {
            null or "aboveAverage" => XLFilterDynamicType.AboveAverage,
            "belowAverage" => XLFilterDynamicType.BelowAverage,
            var value => throw new InvalidDataException($"Unsupported dynamic filter type '{value}'.")
        };
        var threshold = (double)source.Attribute("val")!;
        column.FilterType = XLFilterType.Dynamic;
        column.DynamicType = type;
        column.DynamicValue = threshold;
        column.AddFilter(XLFilter.CreateAverage(threshold, type == XLFilterDynamicType.AboveAverage));
    }

    internal static XElement Write(XLAutoFilter filter)
    {
        var range = filter.Range!;
        var target = new XElement(Spreadsheet + "autoFilter", new XAttribute("ref", range.RangeAddress.ToString()!));
        foreach (var (index, column) in filter.Columns)
        {
            var definition = column.FilterType switch
            {
                XLFilterType.Custom => WriteCustom(column),
                XLFilterType.TopBottom => new XElement(Spreadsheet + "top10",
                    column.TopBottomPart == XLTopBottomPart.Bottom ? new XAttribute("top", "0") : null,
                    column.TopBottomType == XLTopBottomType.Percent ? new XAttribute("percent", "1") : null,
                    new XAttribute("val", column.TopBottomValue)),
                XLFilterType.Dynamic => new XElement(Spreadsheet + "dynamicFilter",
                    new XAttribute("type", column.DynamicType == XLFilterDynamicType.AboveAverage ? "aboveAverage" : "belowAverage"),
                    new XAttribute("val", column.DynamicValue)),
                XLFilterType.Regular => WriteRegular(column),
                _ => throw new NotSupportedException()
            };
            target.Add(new XElement(Spreadsheet + "filterColumn", new XAttribute("colId", index - 1), definition));
        }

        if (filter.Sorted)
        {
            var sortRange = range.RowCount() > 1 ? range.Range(range.FirstCell().CellBelow(), range.LastCell()) : range;
            var columnRange = range.Range(1, filter.SortColumn, range.RowCount(), filter.SortColumn);
            target.Add(new XElement(Spreadsheet + "sortState", new XAttribute("ref", sortRange.RangeAddress.ToString()!),
                new XElement(Spreadsheet + "sortCondition",
                    filter.SortOrder == XLSortOrder.Descending ? new XAttribute("descending", "1") : null,
                    new XAttribute("ref", columnRange.RangeAddress.ToString()!))));
        }

        return target;
    }

    private static XElement WriteCustom(XLFilterColumn column)
    {
        var element = new XElement(Spreadsheet + "customFilters");
        foreach (var filter in column)
        {
            if (filter.Connector == XLConnector.And)
                element.SetAttributeValue("and", "1");

            var op = filter.Operator switch
            {
                XLFilterOperator.Equal => null,
                XLFilterOperator.NotEqual => "notEqual",
                XLFilterOperator.GreaterThan => "greaterThan",
                XLFilterOperator.EqualOrGreaterThan => "greaterThanOrEqual",
                XLFilterOperator.LessThan => "lessThan",
                XLFilterOperator.EqualOrLessThan => "lessThanOrEqual",
                _ => throw new NotSupportedException()
            };
            element.Add(new XElement(Spreadsheet + "customFilter", op is null ? null : new XAttribute("operator", op),
                new XAttribute("val", filter.CustomValue.ToString(CultureInfo.InvariantCulture))));
        }

        return element;
    }

    private static XElement WriteRegular(XLFilterColumn column)
    {
        var element = new XElement(Spreadsheet + "filters");
        foreach (var filter in column)
        {
            if (filter.Value is string text)
                element.Add(new XElement(Spreadsheet + "filter", new XAttribute("val", text)));
        }

        Span<int> date = stackalloc int[6];
        foreach (var filter in column)
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

    private static IXLRange GetRange(XLWorksheet worksheet, XElement element)
    {
        var reference = Required(element, "ref");
        return worksheet.Range(reference) ?? throw new InvalidDataException($"Invalid filter range '{reference}'.");
    }
}
