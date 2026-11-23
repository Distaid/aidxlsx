using System.Xml.Linq;
using System.Xml;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.SpreadsheetML.Tables;

internal static class TableDefinitionXml
{
    internal static void Load(Stream input, XlsxWorksheet worksheet, string relationshipId)
    {
        using var reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        var root = XElement.Load(reader);
        var ns = root.Name.Namespace;
        var reference = (string?)root.Attribute("ref");
        var name = (string?)root.Attribute("name") ?? (string?)root.Attribute("displayName");
        if (string.IsNullOrWhiteSpace(name))
            throw new InvalidDataException("The table name is missing.");

        var range = (reference is null ? null : worksheet.Range(reference)) ?? throw new InvalidDataException($"Invalid table range '{reference}'.");
        var table = (XlsxTable)range.CreateTable(name, false);
        table.RelId = relationshipId;
        var formats = worksheet.Workbook.Styles.DifferentialFormats;

        var columns = root.Element(ns + "tableColumns") ?? throw new InvalidDataException("The table columns are missing.");
        foreach (var column in columns.Elements(ns + "tableColumn"))
        {
            var field = table.AddField(ReadColumnName(column));

            if ((uint?)column.Attribute("headerRowDxfId") is { } header)
                field.HeaderFormatValue = formats[checked((int)header)];
            if ((uint?)column.Attribute("dataDxfId") is { } data)
                field.DataFormatValue = formats[checked((int)data)];
            if ((uint?)column.Attribute("totalsRowDxfId") is { } total)
                field.TotalFormatValue = formats[checked((int)total)];
        }

        if ((uint?)root.Attribute("headerRowCount") == 0)
            table._showHeaderRow = false;
        else
            table.InitializeAutoFilter();

        table._showTotalsRow = (uint?)root.Attribute("totalsRowCount") is > 0;

        if (root.Element(ns + "tableStyleInfo") is { } style)
        {
            table.EmphasizeFirstColumn = (bool?)style.Attribute("showFirstColumn") ?? table.EmphasizeFirstColumn;
            table.EmphasizeLastColumn = (bool?)style.Attribute("showLastColumn") ?? table.EmphasizeLastColumn;
            table.ShowRowStripes = (bool?)style.Attribute("showRowStripes") ?? table.ShowRowStripes;
            table.ShowColumnStripes = (bool?)style.Attribute("showColumnStripes") ?? table.ShowColumnStripes;
            var theme = (string?)style.Attribute("name");
            table.Theme = theme is null ? XlsxTableTheme.None : XlsxTableTheme.FromName(theme) ?? new XlsxTableTheme(theme);
        }

        var filter = root.Element(ns + "autoFilter");
        table.ShowAutoFilter = filter is not null;
        if (filter is not null)
            AutoFilterXml.ReadColumns(filter, table.AutoFilter);

        if (table.ShowTotalsRow)
        {
            foreach (var column in columns.Elements(ns + "tableColumn"))
            {
                var field = table.Field(ReadColumnName(column));
                if ((string?)column.Attribute("totalsRowFunction") is { } function)
                    field.TotalsRowFunction = ParseTotalsFunction(function);
                if (column.Element(ns + "totalsRowFormula") is { } formula)
                    field.TotalsRowFormulaA1 = formula.Value;
                if ((string?)column.Attribute("totalsRowLabel") is { } label)
                    field.TotalsRowLabel = label;
            }
        }

        if (table.AutoFilter is { } autoFilter)
            autoFilter.Range = table.ShowTotalsRow
                ? worksheet.Range(table.RangeAddress.FirstAddress.RowNumber, table.RangeAddress.FirstAddress.ColumnNumber,
                    table.RangeAddress.LastAddress.RowNumber - 1, table.RangeAddress.LastAddress.ColumnNumber)
                : worksheet.Range(table.RangeAddress);
    }

    private static string ReadColumnName(XElement column)
    {
        var name = (string?)column.Attribute("name") ?? throw new InvalidDataException("A table column name is missing.");
        return name.Replace("_x000a_", Environment.NewLine).Replace("_x005f_x000a_", "_x000a_");
    }

    internal static string FormatTotalsFunction(XlsxTotalsRowFunction value) => value switch
    {
        XlsxTotalsRowFunction.None => "none",
        XlsxTotalsRowFunction.Sum => "sum",
        XlsxTotalsRowFunction.Minimum => "min",
        XlsxTotalsRowFunction.Maximum => "max",
        XlsxTotalsRowFunction.Average => "average",
        XlsxTotalsRowFunction.Count => "count",
        XlsxTotalsRowFunction.CountNumbers => "countNums",
        XlsxTotalsRowFunction.StandardDeviation => "stdDev",
        XlsxTotalsRowFunction.Variance => "var",
        XlsxTotalsRowFunction.Custom => "custom",
        _ => throw new ArgumentOutOfRangeException(nameof(value))
    };

    private static XlsxTotalsRowFunction ParseTotalsFunction(string value) => value switch
    {
        "none" => XlsxTotalsRowFunction.None,
        "sum" => XlsxTotalsRowFunction.Sum,
        "min" => XlsxTotalsRowFunction.Minimum,
        "max" => XlsxTotalsRowFunction.Maximum,
        "average" => XlsxTotalsRowFunction.Average,
        "count" => XlsxTotalsRowFunction.Count,
        "countNums" => XlsxTotalsRowFunction.CountNumbers,
        "stdDev" => XlsxTotalsRowFunction.StandardDeviation,
        "var" => XlsxTotalsRowFunction.Variance,
        "custom" => XlsxTotalsRowFunction.Custom,
        _ => throw new InvalidDataException($"Unknown totals row function '{value}'.")
    };
}
