using System.Globalization;
using System.Xml.Linq;
using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.ConditionalFormats;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Opc;
using Aid.Xlsx.Styles.Colors;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Internal.IO;

internal static class ConditionalFormattingXml
{
    private static readonly XNamespace Spreadsheet = OpenXmlConst.Main2006SsNs;
    private static readonly XNamespace Extended = OpenXmlConst.X14Main2009SsNs;
    private static readonly XNamespace Excel = OpenXmlConst.XmMain2006;

    private static readonly (XlsxConditionalFormatType Model, string Xml)[] RuleTypes =
    [
        (XlsxConditionalFormatType.Expression, "expression"), (XlsxConditionalFormatType.CellIs, "cellIs"),
        (XlsxConditionalFormatType.ColorScale, "colorScale"), (XlsxConditionalFormatType.DataBar, "dataBar"),
        (XlsxConditionalFormatType.IconSet, "iconSet"), (XlsxConditionalFormatType.Top10, "top10"),
        (XlsxConditionalFormatType.IsUnique, "uniqueValues"), (XlsxConditionalFormatType.IsDuplicate, "duplicateValues"),
        (XlsxConditionalFormatType.ContainsText, "containsText"), (XlsxConditionalFormatType.NotContainsText, "notContainsText"),
        (XlsxConditionalFormatType.StartsWith, "beginsWith"), (XlsxConditionalFormatType.EndsWith, "endsWith"),
        (XlsxConditionalFormatType.IsBlank, "containsBlanks"), (XlsxConditionalFormatType.NotBlank, "notContainsBlanks"),
        (XlsxConditionalFormatType.IsError, "containsErrors"), (XlsxConditionalFormatType.NotError, "notContainsErrors"),
        (XlsxConditionalFormatType.TimePeriod, "timePeriod"), (XlsxConditionalFormatType.AboveAverage, "aboveAverage")
    ];

    private static readonly (XlsxCFOperator Model, string Xml)[] Operators =
    [
        (XlsxCFOperator.Equal, "equal"), (XlsxCFOperator.NotEqual, "notEqual"),
        (XlsxCFOperator.GreaterThan, "greaterThan"), (XlsxCFOperator.EqualOrGreaterThan, "greaterThanOrEqual"),
        (XlsxCFOperator.LessThan, "lessThan"), (XlsxCFOperator.EqualOrLessThan, "lessThanOrEqual"),
        (XlsxCFOperator.Between, "between"), (XlsxCFOperator.NotBetween, "notBetween"),
        (XlsxCFOperator.Contains, "containsText"), (XlsxCFOperator.NotContains, "notContains"),
        (XlsxCFOperator.StartsWith, "beginsWith"), (XlsxCFOperator.EndsWith, "endsWith")
    ];

    private static readonly (XlsxCFContentType Model, string Xml)[] ThresholdTypes =
    [
        (XlsxCFContentType.Number, "num"), (XlsxCFContentType.Percent, "percent"),
        (XlsxCFContentType.Minimum, "min"), (XlsxCFContentType.Maximum, "max"),
        (XlsxCFContentType.Formula, "formula"), (XlsxCFContentType.Percentile, "percentile")
    ];

    private static readonly (XlsxIconSetStyle Model, string Xml)[] IconSets =
    [
        (XlsxIconSetStyle.ThreeArrows, "3Arrows"), (XlsxIconSetStyle.ThreeArrowsGray, "3ArrowsGray"),
        (XlsxIconSetStyle.ThreeFlags, "3Flags"), (XlsxIconSetStyle.ThreeTrafficLights1, "3TrafficLights1"),
        (XlsxIconSetStyle.ThreeTrafficLights2, "3TrafficLights2"), (XlsxIconSetStyle.ThreeSigns, "3Signs"),
        (XlsxIconSetStyle.ThreeSymbols, "3Symbols"), (XlsxIconSetStyle.ThreeSymbols2, "3Symbols2"),
        (XlsxIconSetStyle.FourArrows, "4Arrows"), (XlsxIconSetStyle.FourArrowsGray, "4ArrowsGray"),
        (XlsxIconSetStyle.FourRedToBlack, "4RedToBlack"), (XlsxIconSetStyle.FourRating, "4Rating"),
        (XlsxIconSetStyle.FourTrafficLights, "4TrafficLights"), (XlsxIconSetStyle.FiveArrows, "5Arrows"),
        (XlsxIconSetStyle.FiveArrowsGray, "5ArrowsGray"), (XlsxIconSetStyle.FiveRating, "5Rating"),
        (XlsxIconSetStyle.FiveQuarters, "5Quarters")
    ];

    private static readonly (XlsxTimePeriod Model, string Xml)[] TimePeriods =
    [
        (XlsxTimePeriod.Yesterday, "yesterday"), (XlsxTimePeriod.Today, "today"), (XlsxTimePeriod.Tomorrow, "tomorrow"),
        (XlsxTimePeriod.InTheLast7Days, "last7Days"), (XlsxTimePeriod.LastWeek, "lastWeek"),
        (XlsxTimePeriod.ThisWeek, "thisWeek"), (XlsxTimePeriod.NextWeek, "nextWeek"),
        (XlsxTimePeriod.LastMonth, "lastMonth"), (XlsxTimePeriod.ThisMonth, "thisMonth"), (XlsxTimePeriod.NextMonth, "nextMonth")
    ];

    internal const string DataBarsExtensionUri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";

    internal static void Load(XElement source, XlsxWorksheet worksheet, LoadContext context)
    {
        var ns = source.Name.Namespace;
        var pivot = (bool?)source.Attribute("pivot") ?? false;
        var references = ((string?)source.Attribute("sqref") ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        foreach (var rule in source.Elements(ns + "cfRule"))
        {
            if ((string?)rule.Attribute("type") is not { } type)
                continue;

            var ranges = references.Select(reference => worksheet.Range(reference)
                ?? throw new InvalidDataException($"Invalid conditional formatting range '{reference}'."));
            var format = new XlsxConditionalFormat(worksheet, XlsxAreaList.FromRanges(worksheet, ranges))
            {
                ConditionalFormatType = ReadValue(RuleTypes, type),
                StopIfTrue = (bool?)rule.Attribute("stopIfTrue") ?? false,
                Priority = (int?)rule.Attribute("priority") ?? int.MaxValue
            };
            if ((uint?)rule.Attribute("dxfId") is { } dxfId)
                format.FormatValue = worksheet.Workbook.Styles.DifferentialFormats[checked((int)dxfId)];

            ReadRuleValues(rule, format);
            if (rule.Element(ns + "colorScale") is { } colorScale)
                ReadThresholds(colorScale, format);
            else if (rule.Element(ns + "dataBar") is { } dataBar)
            {
                format.ShowBarOnly = (bool?)dataBar.Attribute("showValue") == false;
                if (Guid.TryParse(rule.Descendants(Extended + "id").FirstOrDefault()?.Value, out var id))
                    format.Id = id;

                ReadThresholds(dataBar, format);
            }
            else if (rule.Element(ns + "iconSet") is { } iconSet)
            {
                format.ShowIconOnly = (bool?)iconSet.Attribute("showValue") == false;
                format.ReverseIconOrder = (bool?)iconSet.Attribute("reverse") ?? false;
                format.IconSetStyle = ReadValue(IconSets, (string?)iconSet.Attribute("iconSet") ?? "3TrafficLights1");
                ReadThresholds(iconSet, format);
            }

            if (pivot)
                context.AddPivotTableCf(worksheet.Name, format);
            else
                worksheet.ConditionalFormats.Add(format);
        }
    }

    private static void ReadRuleValues(XElement rule, XlsxConditionalFormat format)
    {
        switch (format.ConditionalFormatType)
        {
            case XlsxConditionalFormatType.CellIs:
            case XlsxConditionalFormatType.Expression:
                var formulaCount = 1;
                if (format.ConditionalFormatType == XlsxConditionalFormatType.CellIs)
                {
                    format.Operator = ReadValue(Operators, (string?)rule.Attribute("operator"));
                    if (format.Operator is XlsxCFOperator.Between or XlsxCFOperator.NotBetween)
                        formulaCount = 2;
                }

                // Excel ignores empty formula elements and entries beyond the expected arity.
                var formulas = rule.Elements(rule.Name.Namespace + "formula")
                    .Select(element => element.Value).Where(text => text.Length != 0).Take(formulaCount).ToArray();
                if (formulas.Length != formulaCount)
                    throw PartStructureException.IncorrectElementsCount();

                foreach (var formula in formulas)
                    format.Values.Add(ReadFormula(formula));
                break;
            case XlsxConditionalFormatType.ContainsText:
            case XlsxConditionalFormatType.NotContainsText:
            case XlsxConditionalFormatType.StartsWith:
            case XlsxConditionalFormatType.EndsWith:
                format.Values.Add(new XlsxFormula((string?)rule.Attribute("text") ?? string.Empty) { IsFormula = false });
                break;
            case XlsxConditionalFormatType.Top10:
                format.Percent = (bool?)rule.Attribute("percent") ?? false;
                format.Bottom = (bool?)rule.Attribute("bottom") ?? false;
                if ((uint?)rule.Attribute("rank") is { } rank)
                    format.Values.Add(ReadFormula(rank.ToString(CultureInfo.InvariantCulture)));
                break;
            case XlsxConditionalFormatType.TimePeriod:
                format.TimePeriod = ReadValue(TimePeriods, (string?)rule.Attribute("timePeriod") ?? "yesterday");
                break;
        }
    }

    private static XlsxFormula ReadFormula(string text) => new() { _value = text, IsFormula = !(text.StartsWith('"') && text.EndsWith('"')) };

    private static void ReadThresholds(XElement source, XlsxConditionalFormat format)
    {
        var ns = source.Name.Namespace;
        foreach (var threshold in source.Elements(ns + "cfvo"))
        {
            if ((string?)threshold.Attribute("type") is { } type)
                format.ContentTypes.Add(ReadValue(ThresholdTypes, type));

            var value = (string?)threshold.Attribute("val");
            format.Values.Add(value is null ? null! : new XlsxFormula { Value = value });
            var inclusive = (bool?)threshold.Attribute("gte") != false;
            format.IconSetOperators.Add(inclusive ? XlsxCFIconSetOperator.EqualOrGreaterThan : XlsxCFIconSetOperator.GreaterThan);
        }

        foreach (var color in source.Elements(ns + "color"))
            format.Colors.Add(ReadColor(color));
    }

    internal static void LoadDataBarExtensions(XElement source, XlsxWorksheet worksheet)
    {
        foreach (var rule in source.Descendants(Extended + "cfRule"))
        {
            if ((string?)rule.Attribute("type") != "dataBar" ||
                !Guid.TryParse((string?)rule.Attribute("id"), out var id) ||
                rule.Element(Extended + "dataBar")?.Element(Extended + "negativeFillColor") is not { } negativeFill)
                continue;

            var format = worksheet.ConditionalFormats.SingleOrDefault<XlsxConditionalFormat>(candidate => candidate.Id == id);
            format?.Colors.Add(ReadColor(negativeFill));
        }
    }

    private static XlsxColor ReadColor(XElement color)
    {
        if ((string?)color.Attribute("rgb") is { } rgb)
            return XlsxColor.FromColor(ColorStringParser.ParseFromArgb(rgb.AsSpan()));
        if ((uint?)color.Attribute("indexed") is <= 64 and { } indexed)
            return XlsxColor.FromIndex((int)indexed);
        if ((uint?)color.Attribute("theme") is { } theme)
            return XlsxColor.FromTheme((XlsxThemeColor)theme, (double?)color.Attribute("tint") ?? 0);

        return XlsxColor.Automatic;
    }

    internal static XElement WriteGroup(string references, bool pivot, IEnumerable<XlsxConditionalFormat> formats, XlsxWorkbook.SaveContext context)
    {
        var group = new XElement(Spreadsheet + "conditionalFormatting");
        if (pivot)
            group.Add(Flag("pivot", true));
        group.Add(new XAttribute("sqref", references));
        foreach (var format in formats)
            group.Add(WriteRule(format, context));

        return group;
    }

    private static XElement WriteRule(XlsxConditionalFormat format, XlsxWorkbook.SaveContext context)
    {
        var rule = new XElement(Spreadsheet + "cfRule", new XAttribute("type", WriteValue(RuleTypes, format.ConditionalFormatType)));
        if (format.FormatValue is not null && format.ConditionalFormatType is not (XlsxConditionalFormatType.DataBar or XlsxConditionalFormatType.ColorScale or XlsxConditionalFormatType.IconSet))
            rule.Add(new XAttribute("dxfId", context.GetDxfId(format.FormatValue)));
        rule.Add(new XAttribute("priority", format.Priority));
        if (format.StopIfTrue)
            rule.Add(Flag("stopIfTrue", true));

        switch (format.ConditionalFormatType)
        {
            case XlsxConditionalFormatType.Expression:
            case XlsxConditionalFormatType.CellIs:
                rule.Add(new XAttribute("operator", WriteValue(Operators, format.Operator)));
                AddFormula(rule, QuoteOperand(format.Values[1]));
                if (format.Operator is XlsxCFOperator.Between or XlsxCFOperator.NotBetween)
                    AddFormula(rule, QuoteOperand(format.Values[2]));
                break;
            case XlsxConditionalFormatType.ContainsText:
            case XlsxConditionalFormatType.NotContainsText:
            case XlsxConditionalFormatType.StartsWith:
            case XlsxConditionalFormatType.EndsWith:
                WriteTextRule(rule, format);
                break;
            case XlsxConditionalFormatType.IsBlank:
                AddFormula(rule, $"LEN(TRIM({FirstCell(format)}))=0");
                break;
            case XlsxConditionalFormatType.NotBlank:
                AddFormula(rule, $"LEN(TRIM({FirstCell(format)}))>0");
                break;
            case XlsxConditionalFormatType.IsError:
                AddFormula(rule, $"ISERROR({FirstCell(format)})");
                break;
            case XlsxConditionalFormatType.NotError:
                AddFormula(rule, $"NOT(ISERROR({FirstCell(format)}))");
                break;
            case XlsxConditionalFormatType.Top10:
                rule.Add(Flag("percent", format.Percent), Flag("bottom", format.Bottom),
                    new XAttribute("rank", uint.Parse(format.Values[1].Value, CultureInfo.InvariantCulture)));
                break;
            case XlsxConditionalFormatType.TimePeriod:
                rule.Add(new XAttribute("timePeriod", WriteValue(TimePeriods, format.TimePeriod)));
                AddFormula(rule, DateFormula(format.TimePeriod, FirstCell(format)));
                break;
            case XlsxConditionalFormatType.ColorScale:
                rule.Add(WriteColorScale(format));
                break;
            case XlsxConditionalFormatType.DataBar:
                rule.Add(WriteDataBar(format));
                rule.Add(new XElement(Spreadsheet + "extLst", new XElement(Spreadsheet + "ext",
                    new XAttribute("uri", "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}"), new XAttribute(XNamespace.Xmlns + "x14", Extended),
                    new XElement(Extended + "id", format.Id.WrapInBraces()))));
                break;
            case XlsxConditionalFormatType.IconSet:
                rule.Add(WriteIconSet(format));
                break;
            case XlsxConditionalFormatType.IsUnique:
            case XlsxConditionalFormatType.IsDuplicate:
                break;
            default:
                throw new NotImplementedException($"Conditional formatting rule '{format.ConditionalFormatType}' is not implemented.");
        }

        return rule;
    }

    private static void WriteTextRule(XElement rule, XlsxConditionalFormat format)
    {
        var text = format.Values[1].Value;
        var cell = FirstCell(format);
        var length = text.Length.ToString(CultureInfo.InvariantCulture);
        var (op, formula) = format.ConditionalFormatType switch
        {
            XlsxConditionalFormatType.ContainsText => ("containsText", $"NOT(ISERROR(SEARCH(\"{text}\",{cell})))"),
            XlsxConditionalFormatType.NotContainsText => ("notContains", $"ISERROR(SEARCH(\"{text}\",{cell}))"),
            XlsxConditionalFormatType.StartsWith => ("beginsWith", $"LEFT({cell},{length})=\"{text}\""),
            XlsxConditionalFormatType.EndsWith => ("endsWith", $"RIGHT({cell},{length})=\"{text}\""),
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };
        rule.Add(new XAttribute("operator", op), new XAttribute("text", text));
        AddFormula(rule, formula);
    }

    private static string DateFormula(XlsxTimePeriod period, string cell) => period switch
    {
        XlsxTimePeriod.Today => $"FLOOR({cell},1)=TODAY()",
        XlsxTimePeriod.Yesterday => $"FLOOR({cell},1)=TODAY()-1",
        XlsxTimePeriod.Tomorrow => $"FLOOR({cell},1)=TODAY()+1",
        XlsxTimePeriod.InTheLast7Days => $"AND(TODAY()-FLOOR({cell},1)<=6,FLOOR({cell},1)<=TODAY())",
        XlsxTimePeriod.ThisMonth => $"AND(MONTH({cell})=MONTH(TODAY()),YEAR({cell})=YEAR(TODAY()))",
        XlsxTimePeriod.LastMonth => $"AND(MONTH({cell})=MONTH(EDATE(TODAY(),0-1)),YEAR({cell})=YEAR(EDATE(TODAY(),0-1)))",
        XlsxTimePeriod.NextMonth => $"AND(MONTH({cell})=MONTH(EDATE(TODAY(),0+1)),YEAR({cell})=YEAR(EDATE(TODAY(),0+1)))",
        XlsxTimePeriod.ThisWeek or XlsxTimePeriod.LastWeek => $"AND(TODAY()-ROUNDDOWN({cell},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN({cell},0)-TODAY()<=7-WEEKDAY(TODAY()))",
        XlsxTimePeriod.NextWeek => $"AND(ROUNDDOWN({cell},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN({cell},0)-TODAY()<(15-WEEKDAY(TODAY())))",
        _ => throw new ArgumentOutOfRangeException(nameof(period))
    };

    private static string FirstCell(XlsxConditionalFormat format) => format.Range.RangeAddress.FirstAddress.ToStringRelative(false);

    private static string QuoteOperand(XlsxFormula formula)
    {
        var value = formula.Value;
        if (formula.IsFormula || value.StartsWith('"') && value.EndsWith('"') ||
            double.TryParse(value, XlsxHelper.NumberStyle, XlsxHelper.ParseCulture, out _))
            return value;

        return string.Concat("\"", value.Replace("\"", "\"\""), "\"");
    }

    private static void AddFormula(XElement rule, string formula) => rule.Add(new XElement(Spreadsheet + "formula", formula));

    private static XElement WriteColorScale(XlsxConditionalFormat format)
    {
        var scale = new XElement(Spreadsheet + "colorScale");
        for (var index = 1; index <= format.ContentTypes.Count; index++)
            scale.Add(WriteThreshold(format, index));
        for (var index = 1; index <= format.Colors.Count; index++)
            scale.Add(WriteColor(Spreadsheet + "color", format.Colors[index]));

        return scale;
    }

    private static XElement WriteDataBar(XlsxConditionalFormat format) => new(Spreadsheet + "dataBar",
        Flag("showValue", !format.ShowBarOnly), WriteThreshold(format, 1, XlsxCFContentType.Minimum),
        WriteThreshold(format, 2, XlsxCFContentType.Maximum), WriteColor(Spreadsheet + "color", format.Colors[1]));

    private static XElement WriteIconSet(XlsxConditionalFormat format)
    {
        var set = new XElement(Spreadsheet + "iconSet", new XAttribute("iconSet", WriteValue(IconSets, format.IconSetStyle)),
            Flag("showValue", !format.ShowIconOnly), Flag("reverse", format.ReverseIconOrder));
        for (var index = 1; index <= format.Values.Count; index++)
        {
            var threshold = WriteThreshold(format, index);
            threshold.Add(Flag("gte", format.IconSetOperators[index] == XlsxCFIconSetOperator.EqualOrGreaterThan));
            set.Add(threshold);
        }

        return set;
    }

    private static XElement WriteThreshold(XlsxConditionalFormat format, int index, XlsxCFContentType? fallback = null)
    {
        var type = format.ContentTypes.TryGetValue(index, out var stored) ? stored : fallback!.Value;
        var threshold = new XElement(Spreadsheet + "cfvo", new XAttribute("type", WriteValue(ThresholdTypes, type)));
        if (format.Values.TryGetValue(index, out var value) && value?.Value is { } text)
            threshold.Add(new XAttribute("val", text));

        return threshold;
    }

    private static XElement WriteColor(XName name, XlsxColor color) => new(name, color.ColorType switch
    {
        XlsxColorType.Color => new XAttribute("rgb", color.Color.ToHex()),
        XlsxColorType.Theme => new XAttribute("theme", (uint)color.ThemeColor),
        XlsxColorType.Indexed => new XAttribute("indexed", (uint)color.Indexed),
        _ => null
    });

    internal static void UpdateDataBarExtensions(XElement extension, IEnumerable<XlsxConditionalFormat> formats)
    {
        var container = extension.Element(Extended + "conditionalFormattings");
        if (container is null)
        {
            container = new XElement(Extended + "conditionalFormattings");
            extension.Add(container);
        }

        foreach (var format in formats)
        {
            var id = format.Id.WrapInBraces();
            var existing = container.Descendants(Extended + "cfRule")
                .SingleOrDefault(rule => Guid.TryParse((string?)rule.Attribute("id"), out var ruleId) && ruleId == format.Id);
            if (existing?.Parent is { } oldGroup)
            {
                existing.Remove();
                if (!oldGroup.Elements(Extended + "cfRule").Any())
                    oldGroup.Remove();
            }

            var dataBar = new XElement(Extended + "dataBar", new XAttribute("minLength", 0), new XAttribute("maxLength", 100),
                Flag("showValue", !format.ShowBarOnly), Flag("gradient", true),
                WriteExtendedThreshold(format, 1, XlsxCFContentType.Minimum), WriteExtendedThreshold(format, 2, XlsxCFContentType.Maximum),
                new XElement(Extended + "negativeFillColor", new XAttribute("rgb", format.Colors[format.Colors.Count == 2 ? 2 : 1].Color.ToHex())),
                new XElement(Extended + "axisColor", new XAttribute("rgb", XlsxColor.Black.Color.ToHex())));
            var references = string.Join(" ", format.Ranges.Select(range => range.RangeAddress.ToStringRelative(false)));
            container.Add(new XElement(Extended + "conditionalFormatting", new XAttribute(XNamespace.Xmlns + "xm", Excel),
                new XElement(Extended + "cfRule", new XAttribute("type", "dataBar"), new XAttribute("id", id), dataBar),
                new XElement(Excel + "sqref", references)));
        }
    }

    private static XElement WriteExtendedThreshold(XlsxConditionalFormat format, int index, XlsxCFContentType fallback)
    {
        var type = format.ContentTypes.TryGetValue(index, out var stored) ? stored : fallback;
        if (format.Values.TryGetValue(index, out var value) && value?.Value is { } text)
            return new XElement(Extended + "cfvo", new XAttribute("type", "num"), new XElement(Excel + "f", text));

        var token = type switch
        {
            XlsxCFContentType.Minimum => "autoMin",
            XlsxCFContentType.Maximum => "autoMax",
            _ => WriteValue(ThresholdTypes, type)
        };
        return new XElement(Extended + "cfvo", new XAttribute("type", token));
    }

    private static XAttribute Flag(string name, bool value) => new(name, value ? "1" : "0");

    private static T ReadValue<T>((T Model, string Xml)[] values, string? token) where T : struct, Enum
    {
        foreach (var pair in values)
        {
            if (pair.Xml == token)
                return pair.Model;
        }

        throw new InvalidDataException($"Invalid conditional formatting value '{token}'.");
    }

    private static string WriteValue<T>((T Model, string Xml)[] values, T value) where T : struct, Enum
    {
        foreach (var pair in values)
        {
            if (EqualityComparer<T>.Default.Equals(pair.Model, value))
                return pair.Xml;
        }

        throw new ArgumentOutOfRangeException(nameof(value), value, "Invalid conditional formatting value.");
    }
}
