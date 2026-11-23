using System.Xml.Linq;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.ConditionalFormats;
using Aid.Xlsx.Model.DataValidation;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Sparklines;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.SpreadsheetML.Styles;

namespace Aid.Xlsx.SpreadsheetML.Worksheets;

internal sealed partial class WorksheetXml
{
    private const string SparklineUri = "{05C60535-1F16-4fd2-B633-F4F36F0B64E0}";
    private const string ValidationUri = "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}";

    internal void UpdateConditionalFormats(XlsxWorksheet sheet, XlsxWorkbook.SaveContext context)
    {
        var pivotFormats = sheet.PivotTables.SelectMany<XlsxPivotTable, XlsxConditionalFormat>(pivot => pivot.ConditionalFormats.Select(format => format.Format)).ToHashSet();
        XlsxConditionalFormat[] formats = [.. sheet.ConditionalFormats.Cast<XlsxConditionalFormat>().Concat(pivotFormats).OrderBy(format => format.Priority)];
        for (var i = 0; i < formats.Length; i++)
            formats[i].Priority = i + 1;

        Replace("conditionalFormatting", formats.GroupBy(format => (
            References: string.Join(" ", format.Ranges.Select(range => range.RangeAddress.ToStringRelative(false))),
            IsPivot: pivotFormats.Contains(format)))
            .Select(group => ConditionalFormattingXml.WriteGroup(group.Key.References, group.Key.IsPivot, [.. group], context)));

        XlsxConditionalFormat[] dataBars = [.. sheet.ConditionalFormats.Cast<XlsxConditionalFormat>().Where(format => format.ConditionalFormatType == XlsxConditionalFormatType.DataBar)];
        if (dataBars.Length == 0)
        {
            RemoveExtensionPayload(ConditionalFormattingXml.DataBarsExtensionUri, Extended + "conditionalFormattings");
            return;
        }

        ConditionalFormattingXml.UpdateDataBarExtensions(GetOrAddExtension(ConditionalFormattingXml.DataBarsExtensionUri), dataBars);
    }

    internal void UpdateSparklines(XlsxWorksheet sheet)
    {
        XlsxSparklineGroup[] groups = [.. sheet.SparklineGroupsInternal.Where<XlsxSparklineGroup>(group => group.Sparklines.Any())];
        if (groups.Length == 0)
        {
            RemoveExtensionPayload(SparklineUri, Extended + "sparklineGroups");
            return;
        }

        XElement container = UpdateExtensionPayload(SparklineUri, Extended + "sparklineGroups");
        foreach (XlsxSparklineGroup? group in groups)
        {
            XNamespace revision = "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2";
            var element = new XElement(Extended + "sparklineGroup",
                new XAttribute(XNamespace.Xmlns + "xr2", revision),
                new XAttribute(revision + "uid", "{A98FF5F8-AE60-43B5-8001-AD89004F45D3}"),
                new XAttribute("lineWeight", group.LineWeight),
                new XAttribute("type", group.Type switch
                {
                    XlsxSparklineType.Line => "line",
                    XlsxSparklineType.Column => "column",
                    XlsxSparklineType.Stacked => "stacked",
                    _ => throw new ArgumentOutOfRangeException(nameof(sheet)),
                }),
                new XAttribute("displayEmptyCellsAs", group.DisplayEmptyCellsAs switch
                {
                    XlsxDisplayBlanksAsValues.Interpolate => "span",
                    XlsxDisplayBlanksAsValues.NotPlotted => "gap",
                    XlsxDisplayBlanksAsValues.Zero => "zero",
                    _ => throw new ArgumentOutOfRangeException(nameof(sheet)),
                }));
            Flag(element, "high", group.ShowMarkers.HasFlag(XlsxSparklineMarkers.HighPoint));
            Flag(element, "low", group.ShowMarkers.HasFlag(XlsxSparklineMarkers.LowPoint));
            Flag(element, "first", group.ShowMarkers.HasFlag(XlsxSparklineMarkers.FirstPoint));
            Flag(element, "last", group.ShowMarkers.HasFlag(XlsxSparklineMarkers.LastPoint));
            Flag(element, "negative", group.ShowMarkers.HasFlag(XlsxSparklineMarkers.NegativePoints));
            Flag(element, "markers", group.ShowMarkers.HasFlag(XlsxSparklineMarkers.Markers));
            Flag(element, "displayHidden", group.DisplayHidden);
            Flag(element, "displayXAxis", group.HorizontalAxis.IsVisible);
            Flag(element, "rightToLeft", group.HorizontalAxis.RightToLeft);
            Flag(element, "dateAxis", group.HorizontalAxis.DateAxis);

            IXlsxSparklineVerticalAxis axis = group.VerticalAxis;
            element.SetAttributeValue("minAxisType", AxisType(axis.MinAxisType));
            element.SetAttributeValue("maxAxisType", AxisType(axis.MaxAxisType));
            if (axis.MinAxisType == XlsxSparklineAxisMinMax.Custom)
                element.SetAttributeValue("manualMin", axis.ManualMin);
            if (axis.MaxAxisType == XlsxSparklineAxisMinMax.Custom)
                element.SetAttributeValue("manualMax", axis.ManualMax);

            element.Add(SpreadsheetColorXml.Write(Extended + "colorSeries", group.Style.SeriesColor),
                SpreadsheetColorXml.Write(Extended + "colorNegative", group.Style.NegativeColor),
                new XElement(Extended + "colorAxis", new XAttribute("rgb", group.HorizontalAxis.Color.Color.ToHex())),
                SpreadsheetColorXml.Write(Extended + "colorMarkers", group.Style.MarkersColor),
                SpreadsheetColorXml.Write(Extended + "colorFirst", group.Style.FirstMarkerColor),
                SpreadsheetColorXml.Write(Extended + "colorLast", group.Style.LastMarkerColor),
                SpreadsheetColorXml.Write(Extended + "colorHigh", group.Style.HighMarkerColor),
                SpreadsheetColorXml.Write(Extended + "colorLow", group.Style.LowMarkerColor));
            if (group.HorizontalAxis.DateAxis && group.DateRange is { } dates)
                element.Add(new XElement(Excel + "f", dates.RangeAddress.ToString(XlsxReferenceStyle.A1, true)));

            element.Add(new XElement(Extended + "sparklines", group.Sparklines.Select(sparkline => new XElement(Extended + "sparkline",
                // Excel uses an empty formula when the data range has been deleted.
                new XElement(Excel + "f", sparkline.SourceDataFormula ?? string.Empty),
                new XElement(Excel + "sqref", sparkline.Location.ToString())))));
            container.Add(element);
        }

        static string AxisType(XlsxSparklineAxisMinMax value) => value switch
        {
            XlsxSparklineAxisMinMax.Automatic => "individual",
            XlsxSparklineAxisMinMax.SameForAll => "group",
            XlsxSparklineAxisMinMax.Custom => "custom",
            _ => throw new ArgumentOutOfRangeException(nameof(value)),
        };
    }

    internal void UpdateValidations(XlsxWorksheet sheet, bool consolidate)
    {
        if (consolidate)
            sheet.DataValidations.Consolidate();

        var standard = new List<XElement>();
        var extended = new List<XElement>();
        var hasStandardValidation = false;
        foreach (XlsxDataValidation rule in sheet.DataValidations)
        {
            (bool minExternal, string? min) = ResolveFormula(rule.MinValue);
            (bool maxExternal, string? max) = ResolveFormula(rule.MaxValue);
            if (minExternal || maxExternal)
            {
                extended.Add(WriteValidation(rule, min, max, Extended));
            }
            else
            {
                standard.Add(WriteValidation(rule, min, max, Main));
                hasStandardValidation |= rule.IsDirty();
            }
        }

        if (hasStandardValidation)
        {
            XElement element = GetOrAdd("dataValidations");
            element.Elements(Main + "dataValidation").Remove();
            element.Add(standard);
            element.SetAttributeValue("count", standard.Count);
        }
        else
        {
            Replace("dataValidations", []);
        }

        if (extended.Count > 0)
        {
            XElement element = UpdateExtensionPayload(ValidationUri, Extended + "dataValidations");
            element.Add(extended);
            element.SetAttributeValue("count", extended.Count);
        }
        else
        {
            RemoveExtensionPayload(ValidationUri, Extended + "dataValidations");
        }

        (bool External, string Formula) ResolveFormula(string formula)
        {
            if (!XlsxHelper.IsValidRangeAddress(formula))
                return (false, formula);

            var separator = formula.LastIndexOf('!');
            if (separator < 0)
                return (false, formula);

            // Same-sheet range references are stored without the sheet name.
            return XlsxHelper.SheetComparer.Equals(sheet.Name, formula[..separator].UnescapeSheetName())
                ? (false, formula[(separator + 1)..])
                : (true, formula);
        }
    }

    private static XElement WriteValidation(IXlsxDataValidation rule, string min, string max, XNamespace ns)
    {
        var element = new XElement(ns + "dataValidation");
        Flag(element, "allowBlank", rule.IgnoreBlanks);
        Flag(element, "showErrorMessage", rule.ShowErrorMessage);
        Flag(element, "showDropDown", !rule.InCellDropdown);
        Flag(element, "showInputMessage", rule.ShowInputMessage);
        element.SetAttributeValue("prompt", rule.InputMessage);
        element.SetAttributeValue("promptTitle", rule.InputTitle);
        element.SetAttributeValue("errorTitle", rule.ErrorTitle);
        element.SetAttributeValue("error", rule.ErrorMessage);
        element.SetAttributeValue("type", rule.AllowedValues switch
        {
            XlsxAllowedValues.AnyValue => "none",
            XlsxAllowedValues.WholeNumber => "whole",
            XlsxAllowedValues.Decimal => "decimal",
            XlsxAllowedValues.List => "list",
            XlsxAllowedValues.Date => "date",
            XlsxAllowedValues.Time => "time",
            XlsxAllowedValues.TextLength => "textLength",
            XlsxAllowedValues.Custom => "custom",
            _ => throw new ArgumentOutOfRangeException(nameof(rule)),
        });
        element.SetAttributeValue("errorStyle", rule.ErrorStyle switch
        {
            XlsxErrorStyle.Stop => "stop",
            XlsxErrorStyle.Warning => "warning",
            XlsxErrorStyle.Information => "information",
            _ => throw new ArgumentOutOfRangeException(nameof(rule)),
        });
        element.SetAttributeValue("operator", rule.Operator switch
        {
            XlsxOperator.Between => "between",
            XlsxOperator.NotBetween => "notBetween",
            XlsxOperator.EqualTo => "equal",
            XlsxOperator.NotEqualTo => "notEqual",
            XlsxOperator.GreaterThan => "greaterThan",
            XlsxOperator.LessThan => "lessThan",
            XlsxOperator.EqualOrGreaterThan => "greaterThanOrEqual",
            XlsxOperator.EqualOrLessThan => "lessThanOrEqual",
            _ => throw new ArgumentOutOfRangeException(nameof(rule)),
        });

        var references = string.Join(" ", rule.Ranges.Select(range => range.RangeAddress));
        if (ns == Main)
        {
            element.SetAttributeValue("sqref", references);
            element.Add(new XElement(ns + "formula1", min), new XElement(ns + "formula2", max));
        }
        else
        {
            if (!string.IsNullOrWhiteSpace(min))
                element.Add(new XElement(ns + "formula1", new XElement(Excel + "f", min)));
            if (!string.IsNullOrWhiteSpace(max))
                element.Add(new XElement(ns + "formula2", new XElement(Excel + "f", max)));
            element.Add(new XElement(Excel + "sqref", references));
        }

        return element;
    }

    private XElement? FindExtension(string uri) => Root.Element(Main + "extLst")?.Elements(Main + "ext")
        .FirstOrDefault(extension => string.Equals((string?)extension.Attribute("uri"), uri, StringComparison.OrdinalIgnoreCase));

    private XElement GetOrAddExtension(string uri)
    {
        if (FindExtension(uri) is { } existing)
            return existing;

        var extension = new XElement(Main + "ext", new XAttribute("uri", uri), new XAttribute(XNamespace.Xmlns + "x14", Extended));
        GetOrAdd("extLst").Add(extension);
        return extension;
    }

    private XElement UpdateExtensionPayload(string uri, XName name)
    {
        XElement extension = GetOrAddExtension(uri);
        XElement? payload = extension.Element(name);
        if (payload is null)
        {
            payload = new XElement(name, new XAttribute(XNamespace.Xmlns + "xm", Excel));
            extension.Add(payload);
        }
        else
        {
            payload.RemoveNodes();
        }

        return payload;
    }

    private void RemoveExtensionPayload(string uri, XName name)
    {
        if (FindExtension(uri) is not { } extension)
            return;

        extension.Elements(name).Remove();
        if (extension.Nodes().Any())
            return;

        XElement extensions = extension.Parent!;
        extension.Remove();
        if (!extensions.Nodes().Any())
            extensions.Remove();
    }
}
