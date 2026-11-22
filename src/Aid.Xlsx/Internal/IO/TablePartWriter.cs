using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using static Aid.Xlsx.XLWorkbook;

namespace Aid.Xlsx.Internal.IO
{
    /// <summary>
    /// A writer for table definition part.
    /// </summary>
    internal class TablePartWriter
    {
        internal static void SynchronizeTableParts(XLTables tables, WorksheetPart worksheetPart, SaveContext context)
        {
            // Remove table definition parts that are not a part of workbook
            foreach (var tableDefinitionPart in worksheetPart.GetPartsOfType<TableDefinitionPart>().ToList())
            {
                var partId = worksheetPart.GetIdOfPart(tableDefinitionPart);
                var xlWorkbookContainsTable = tables.Cast<XLTable>().Any(t => t.RelId == partId);
                if (!xlWorkbookContainsTable)
                {
                    worksheetPart.DeletePart(tableDefinitionPart);
                }
            }

            foreach (var xlTable in tables.Cast<XLTable>())
            {
                if (String.IsNullOrEmpty(xlTable.RelId))
                {
                    xlTable.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    worksheetPart.AddNewPart<TableDefinitionPart>(xlTable.RelId);
                }
            }
        }

        internal static void GenerateTableParts(XLTables tables, WorksheetPart worksheetPart, SaveContext context)
        {
            foreach (var xlTable in tables.Cast<XLTable>())
            {
                var relId = xlTable.RelId;
                var tableDefinitionPart = (TableDefinitionPart)worksheetPart.GetPartById(relId);
                using var output = tableDefinitionPart.GetStream(FileMode.Create, FileAccess.Write);
                Write(output, xlTable, context);
            }
        }

        private static void Write(Stream output, XLTable xlTable, SaveContext context)
        {
            XNamespace ns = OpenXmlConst.Main2006SsNs;
            context.TableId++;
            var reference = xlTable.RangeAddress.FirstAddress + ":" + xlTable.RangeAddress.LastAddress;
            var tableName = GetTableName(xlTable.Name, context);
            var table = new XElement(ns + "table", new XAttribute(XNamespace.Xmlns + "x", ns),
                new XAttribute("id", context.TableId), new XAttribute("name", tableName),
                new XAttribute("displayName", tableName), new XAttribute("ref", reference));

            if (!xlTable.ShowHeaderRow)
                table.SetAttributeValue("headerRowCount", 0);

            if (xlTable.ShowTotalsRow)
                table.SetAttributeValue("totalsRowCount", 1);
            else
                table.SetAttributeValue("totalsRowShown", "0");

            var tableColumns = new XElement(ns + "tableColumns", new XAttribute("count", xlTable.ColumnCount()));

            uint columnId = 0;
            foreach (var xlField in xlTable.Fields)
            {
                columnId++;
                var fieldName = xlField.Name;
                var tableColumn = new XElement(ns + "tableColumn", new XAttribute("id", columnId),
                    new XAttribute("name", fieldName.Replace("_x000a_", "_x005f_x000a_").Replace(Environment.NewLine, "_x000a_")));

                // OI-29500: the headerRowDxfId attribute shall be absent if the enclosing table element's headerRow
                // attribute value is greater than or equal to 1. The header dxf stores the format while the table
                // doesn't have a header, but user might want to add it back.
                if (xlField.HeaderFormatValue is { } headerDfx)
                    tableColumn.SetAttributeValue("headerRowDxfId", context.GetDxfId(headerDfx));

                if (xlField.DataFormatValue is { } dataDfx)
                    tableColumn.SetAttributeValue("dataDxfId", context.GetDxfId(dataDfx));

                if (xlField.TotalFormatValue is { } totalsDfx)
                    tableColumn.SetAttributeValue("totalsRowDxfId", context.GetDxfId(totalsDfx));

                if (xlField.IsConsistentFormula())
                {
                    string formula = xlField.Column.Cells()
                        .Skip(xlTable.ShowHeaderRow ? 1 : 0)
                        .First()
                        .FormulaA1;

                    while (formula.StartsWith('=') && formula.Length > 1)
                        formula = formula.Substring(1);

                    if (!String.IsNullOrWhiteSpace(formula))
                    {
                        tableColumn.Add(new XElement(ns + "calculatedColumnFormula", formula));
                    }
                }

                if (xlTable.ShowTotalsRow)
                {
                    if (xlField.TotalsRowFunction != XLTotalsRowFunction.None)
                    {
                        tableColumn.SetAttributeValue("totalsRowFunction", TableDefinitionXml.FormatTotalsFunction(xlField.TotalsRowFunction));

                        if (xlField.TotalsRowFunction == XLTotalsRowFunction.Custom)
                            tableColumn.Add(new XElement(ns + "totalsRowFormula", xlField.TotalsRowFormulaA1));
                    }

                    if (!String.IsNullOrWhiteSpace(xlField.TotalsRowLabel))
                        tableColumn.SetAttributeValue("totalsRowLabel", xlField.TotalsRowLabel);
                }
                tableColumns.Add(tableColumn);
            }

            var tableStyleInfo = new XElement(ns + "tableStyleInfo",
                new XAttribute("showFirstColumn", xlTable.EmphasizeFirstColumn ? "1" : "0"),
                new XAttribute("showLastColumn", xlTable.EmphasizeLastColumn ? "1" : "0"),
                new XAttribute("showRowStripes", xlTable.ShowRowStripes ? "1" : "0"),
                new XAttribute("showColumnStripes", xlTable.ShowColumnStripes ? "1" : "0"));

            if (xlTable.Theme != XLTableTheme.None)
                tableStyleInfo.SetAttributeValue("name", xlTable.Theme.Name);

            if (xlTable.ShowAutoFilter)
            {
                if (xlTable.ShowTotalsRow)
                {
                    xlTable.AutoFilter.Range = xlTable.Worksheet.Range(
                        xlTable.RangeAddress.FirstAddress.RowNumber, xlTable.RangeAddress.FirstAddress.ColumnNumber,
                        xlTable.RangeAddress.LastAddress.RowNumber - 1, xlTable.RangeAddress.LastAddress.ColumnNumber);
                }
                else
                    xlTable.AutoFilter.Range = xlTable.Worksheet.Range(xlTable.RangeAddress);

                table.Add(AutoFilterXml.Write(xlTable.AutoFilter));
            }

            table.Add(tableColumns, tableStyleInfo);
            new XDocument(table).Save(output, System.Xml.Linq.SaveOptions.DisableFormatting);
        }

        private static string GetTableName(String originalTableName, SaveContext context)
        {
            var tableName = originalTableName.RemoveSpecialCharacters();
            var name = tableName;
            if (context.TableNames.Contains(name))
            {
                var i = 1;
                name = tableName + i.ToInvariantString();
                while (context.TableNames.Contains(name))
                {
                    i++;
                    name = tableName + i.ToInvariantString();
                }
            }

            context.TableNames.Add(name);
            return name;
        }
    }
}
