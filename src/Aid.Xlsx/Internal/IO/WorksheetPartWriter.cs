#nullable disable

using System.Text;
using System.Xml;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Drawings;
using Aid.Xlsx.Exceptions;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Rows;
using Aid.Xlsx.Tables;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;
using DocumentFormat.OpenXml.Packaging;
using XElement = System.Xml.Linq.XElement;
using static Aid.Xlsx.Internal.IO.OpenXmlConst;
using static Aid.Xlsx.Workbooks.XlsxWorkbook;

namespace Aid.Xlsx.Internal.IO
{
    internal class WorksheetPartWriter
    {
        internal static void GenerateWorksheetPartContent(
            bool partIsEmpty,
            WorksheetPart worksheetPart,
            XlsxWorksheet xlWorksheet,
            SaveOptions options,
            SaveContext context)
        {
            if (options.ConsolidateConditionalFormatRanges)
                xlWorksheet.ConditionalFormats.Consolidate();

            WorksheetXml xml;
            using (var source = partIsEmpty ? null : worksheetPart.GetStream(FileMode.Open, FileAccess.Read))
                xml = WorksheetXml.Create(source, xlWorksheet.Internals.RowsCollection.Values.Any(row => row.DyDescent.HasValue));

            xml.UpdateSettings(xlWorksheet);
            xml.UpdateColumns(xlWorksheet, context);
            xml.UpdateConditionalFormats(xlWorksheet, context);
            xml.UpdateSparklines(xlWorksheet);
            xml.UpdateValidations(xlWorksheet, options.ConsolidateDataValidationRanges);
            xml.UpdateHyperlinks(xlWorksheet, worksheetPart.HyperlinkRelationships.ToList().Select(rel => rel.Id),
                id => worksheetPart.DeleteReferenceRelationship(id), uri =>
                {
                    var id = context.RelIdGenerator.GetNext(RelType.Workbook);
                    worksheetPart.AddHyperlinkRelationship(uri, true, id);
                    return id;
                });
            xml.UpdatePrintSettings(xlWorksheet);

            var tables = (XlsxTables)xlWorksheet.Tables;
            var emptyTable = tables.FirstOrDefault<XlsxTable>(table => table.DataRange is null);
            if (emptyTable is not null)
                throw new EmptyTableException($"Table '{emptyTable.Name}' should have at least 1 row.");
            xml.UpdateTableReferences(tables.Cast<XlsxTable>().Select(table => table.RelId));
            tables.Deleted.Clear();

            var previousDrawing = worksheetPart.DrawingsPart;
            var previousDrawingId = previousDrawing is null ? null : worksheetPart.GetIdOfPart(previousDrawing);
            var drawingId = SavePictures(worksheetPart, (XlsxPictures)xlWorksheet.Pictures, context);
            xml.UpdateDrawingReferences(previousDrawingId, drawingId, xlWorksheet.LegacyDrawingId);

            using var stream = worksheetPart.GetStream(FileMode.Create, FileAccess.Write);
            using var writer = XmlWriter.Create(stream, new XmlWriterSettings { Encoding = new UTF8Encoding(false) });
            xml.Write(writer, () => StreamSheetData(writer, xlWorksheet, context, options));
        }

        private static void WriteCellValue(XmlWriter w, XlsxCell xlCell, SaveContext context)
        {
            var dataType = xlCell.DataType;
            if (dataType == XlsxDataType.Blank)
                return;

            if (dataType == XlsxDataType.Text)
            {
                var text = xlCell.GetText();
                if (xlCell.HasFormula)
                {
                    WriteStringValue(w, text);
                }
                else
                {
                    if (xlCell.ShareString)
                    {
                        var sharedStringId = context.GetSharedStringId(xlCell, text);
                        w.WriteStartElement("v", Main2006SsNs);
                        w.WriteValue(sharedStringId);
                        w.WriteEndElement();
                    }
                    else
                    {
                        w.WriteStartElement("is", Main2006SsNs);
                        var richText = xlCell.RichText;
                        if (richText is not null)
                        {
                            TextSerializer.WriteRichTextElements(w, richText, context);
                        }
                        else
                        {
                            w.WriteStartElement("t", Main2006SsNs);
                            if (text.PreserveSpaces())
                                w.WritePreserveSpaceAttr();

                            w.WriteString(text);
                            w.WriteEndElement();
                        }

                        w.WriteEndElement(); // is
                    }
                }
            }
            else if (dataType == XlsxDataType.TimeSpan)
            {
                WriteNumberValue(w, xlCell.Value.GetUnifiedNumber());
            }
            else if (dataType == XlsxDataType.Number)
            {
                WriteNumberValue(w, xlCell.Value.GetNumber());
            }
            else if (dataType == XlsxDataType.DateTime)
            {
                // OpenXML SDK validator requires a specific format, in addition to the spec, but can reads many more
                var date = xlCell.GetDateTime();
                if (xlCell.Worksheet.Workbook.Use1904DateSystem)
                    date = date.AddDays(-1462);

                WriteNumberValue(w, date.ToSerialDateTime());
            }
            else if (dataType == XlsxDataType.Boolean)
            {
                WriteStringValue(w, xlCell.GetBoolean() ? TrueValue : FalseValue);
            }
            else if (dataType == XlsxDataType.Error)
            {
                WriteStringValue(w, xlCell.Value.GetError().ToDisplayString());
            }
            else
            {
                throw new InvalidOperationException();
            }

            static void WriteStringValue(XmlWriter w, String text)
            {
                w.WriteStartElement("v", Main2006SsNs);
                w.WriteString(text);
                w.WriteEndElement();
            }

            static void WriteNumberValue(XmlWriter w, Double value)
            {
                w.WriteStartElement("v", Main2006SsNs);
                w.WriteNumberValue(value);
                w.WriteEndElement();
            }
        }

        private static string SavePictures(WorksheetPart worksheetPart, XlsxPictures pictures, SaveContext context)
        {
            var part = worksheetPart.DrawingsPart;
            if (pictures.Count == 0 && pictures.Deleted.Count == 0)
                return part is null ? null : worksheetPart.GetIdOfPart(part);

            XElement root;
            if (part is null)
            {
                part = worksheetPart.AddNewPart<DrawingsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                root = DrawingXml.Create();
            }
            else
            {
                using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
                root = stream.Length == 0 ? DrawingXml.Create() : DrawingXml.Read(stream);
            }

            var anchorsByImage = root.Elements().Where(DrawingXml.IsPictureAnchor)
                .Select(anchor => (Anchor: anchor, Id: DrawingXml.ImageId(anchor)))
                .Where(item => item.Id is not null).ToLookup(item => item.Id, item => item.Anchor, StringComparer.Ordinal);
            var matched = new HashSet<XElement>();
            foreach (var picture in pictures.Cast<XlsxPicture>())
            {
                ImagePart imagePart;
                if (picture.RelId is not null && part.TryGetPartById(picture.RelId, out var related) && related is ImagePart image)
                {
                    imagePart = image;
                }
                else
                {
                    picture.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    imagePart = part.AddImagePart(picture.Format.ToOpenXml(), picture.RelId);
                }

                picture.ImageStream.Position = 0;
                imagePart.FeedData(picture.ImageStream);

                var anchor = DrawingXml.CreatePictureAnchor(root.Name.Namespace, picture, picture.RelId);
                var candidates = anchorsByImage[picture.RelId].Where(element => !matched.Contains(element)).ToList();
                var existing = candidates.FirstOrDefault(element => element.Descendants(root.Name.Namespace + "cNvPr")
                    .Any(properties => (string)properties.Attribute("name") == picture.Name));
                existing ??= candidates.FirstOrDefault(element => element.Descendants(root.Name.Namespace + "cNvPr")
                    .Any(properties => (int?)properties.Attribute("id") == picture.Id));
                if (existing is null && candidates.Count == 1 && !pictures.Deleted.Contains(picture.RelId))
                    existing = candidates[0];

                if (existing is null)
                {
                    root.Add(anchor);
                }
                else
                {
                    matched.Add(existing);
                    existing.ReplaceWith(anchor);
                }
            }

            foreach (var removedId in pictures.Deleted.Distinct())
            {
                foreach (var anchor in anchorsByImage[removedId].Where(element => !matched.Contains(element)))
                    anchor.Remove();

                if (!root.Elements().Any(element => DrawingXml.ImageId(element) == removedId) && part.HasPartWithId(removedId))
                    part.DeletePart(removedId);
            }

            pictures.Deleted.Clear();
            if (!root.HasElements && !part.Parts.Any())
            {
                worksheetPart.DeletePart(part);
                return null;
            }

            if (pictures.Count != 0)
            {
                var id = 1;
                foreach (var properties in root.Descendants(root.Name.Namespace + "cNvPr"))
                    properties.SetAttributeValue("id", id++);
            }

            using var output = part.GetStream(FileMode.Create, FileAccess.Write);
            DrawingXml.Write(output, root);
            return worksheetPart.GetIdOfPart(part);
        }

        private static void StreamSheetData(XmlWriter xml, XlsxWorksheet xlWorksheet, SaveContext context, SaveOptions options)
        {
            var maxColumn = GetMaxColumn(xlWorksheet);

            xml.WriteStartElement("sheetData", Main2006SsNs);

            var tableTotalCells = new HashSet<IXlsxAddress>(
                xlWorksheet.Tables
                .Where<XlsxTable>(table => table.ShowTotalsRow)
                .SelectMany(table =>
                    table.TotalsRow().CellsUsed())
                .Select(cell => cell.Address));

            // A rather complicated state machine, so rows and cells can be written in a single loop
            var openedRowNumber = 0;
            var isRowOpened = false;
            var cellRef = new char[10]; // Buffer, must be enough to hold span and rowNumber as strings
            var rows = xlWorksheet.Internals.RowsCollection.Keys.ToList();
            rows.Sort();
            var rowPropIndex = 0;
            uint rowStyleId = 0;
            foreach (var xlCell in xlWorksheet.Internals.CellsCollection.GetCells())
            {
                var currentRowNumber = xlCell.Point.Row;

                // A space between cells can have several rows that don't contain cells,
                // but have custom properties (e.g. height). Write them out.
                while (rowPropIndex < rows.Count && rows[rowPropIndex] < currentRowNumber)
                {
                    if (isRowOpened)
                    {
                        xml.WriteEndElement(); // row
                        isRowOpened = false;
                    }

                    var rowNumber = rows[rowPropIndex];
                    var xlRow = xlWorksheet.Internals.RowsCollection[rowNumber];
                    if (RowHasCustomProps(xlRow))
                    {
                        WriteStartRow(xml, xlRow, rowNumber, maxColumn, context);

                        isRowOpened = true;
                        openedRowNumber = rowNumber;
                    }

                    rowPropIndex++;
                }

                // For saving cells to file, ignore conditional formatting, data validation rules and merged
                // ranges. They just bloat the file
                var isEmpty = xlCell.CachedValue.Type == XlsxDataType.Blank &&
                              xlCell.IsEmpty(XlsxCellsUsedOptions.All
                                             & ~XlsxCellsUsedOptions.ConditionalFormats
                                             & ~XlsxCellsUsedOptions.DataValidation
                                             & ~XlsxCellsUsedOptions.MergedRanges);

                if (isEmpty)
                    continue;

                if (openedRowNumber != currentRowNumber)
                {
                    if (isRowOpened)
                        xml.WriteEndElement(); // row

                    if (xlWorksheet.Internals.RowsCollection.TryGetValue(currentRowNumber, out var row))
                    {
                        rowPropIndex++;
                        rowStyleId = context.GetStyleId(row.FormatValue);
                    }
                    else
                    {
                        rowStyleId = 0;
                    }

                    WriteStartRow(xml, row, currentRowNumber, maxColumn, context);

                    isRowOpened = true;
                    openedRowNumber = currentRowNumber;
                }

                WriteCell(xml, xlCell, cellRef, context, options, tableTotalCells, rowStyleId);
            }

            if (isRowOpened)
                xml.WriteEndElement(); // row

            // Write rows with custom properties after last cell.
            while (rowPropIndex < rows.Count)
            {
                var rowNumber = rows[rowPropIndex];
                var xlRow = xlWorksheet.Internals.RowsCollection[rowNumber];
                if (RowHasCustomProps(xlRow))
                {
                    WriteStartRow(xml, xlRow, rowNumber, 0, context);
                    xml.WriteEndElement(); // row
                }

                rowPropIndex++;
            }

            xml.WriteEndElement(); // SheetData

            static bool RowHasCustomProps(XlsxRow xlRow)
            {
                return xlRow.HeightChanged ||
                    xlRow.IsHidden ||
                    xlRow.FormatValue is not null ||
                    xlRow.Collapsed ||
                    xlRow.OutlineLevel > 0;
            }

            static void WriteStartRow(XmlWriter w, XlsxRow xlRow, int rowNumber, int maxColumn, SaveContext context)
            {
                w.WriteStartElement("row", Main2006SsNs);

                w.WriteStartAttribute("r");
                w.WriteValue(rowNumber);
                w.WriteEndAttribute();

                if (maxColumn > 0)
                {
                    w.WriteStartAttribute("spans");
                    w.WriteString("1:");
                    w.WriteValue(maxColumn);
                    w.WriteEndAttribute();
                }

                if (xlRow is null)
                    return;

                if (xlRow.HeightChanged)
                {
                    var height = xlRow.Height.SaveRound();
                    w.WriteStartAttribute("ht");
                    w.WriteNumberValue(height);
                    w.WriteEndAttribute();

                    // Note that dyDescent automatically implies custom height
                    w.WriteAttributeString("customHeight", TrueValue);
                }

                if (xlRow.IsHidden)
                {
                    w.WriteAttributeString("hidden", TrueValue);
                }

                var rowHasCustomFormat = xlRow.FormatValue is not null;
                if (rowHasCustomFormat)
                {
                    var formatIndex = context.GetStyleId(xlRow.FormatValue);
                    w.WriteAttribute("s", formatIndex);
                    w.WriteAttributeString("customFormat", TrueValue);
                }

                if (xlRow.Collapsed)
                {
                    w.WriteAttributeString("collapsed", TrueValue);
                }

                if (xlRow.OutlineLevel > 0)
                {
                    w.WriteAttribute("outlineLevel", xlRow.OutlineLevel);
                }

                if (xlRow.ShowPhonetic)
                {
                    w.WriteAttributeString("ph", TrueValue);
                }

                if (xlRow.DyDescent is not null)
                {
                    w.WriteAttribute("dyDescent", X14Ac2009SsNs, xlRow.DyDescent.Value);
                }

                // thickBot and thickTop attributes are not written, because Excel seems to determine adjustments
                // from cell borders on its own and it would be rather costly to check each cell in each row.
                // If row was adjusted when cell had it's border modified, then it would be fine to write
                // the thickBot/thickBot attributes.
            }

            static void WriteStartCell(XmlWriter w, XlsxCell xlCell, Char[] reference, int referenceLength, String dataType, UInt32 styleId)
            {
                w.WriteStartElement("c", Main2006SsNs);

                w.WriteStartAttribute("r");
                w.WriteRaw(reference, 0, referenceLength);
                w.WriteEndAttribute();

                // TODO: if (styleId != 0) Test files have style even for 0, fix later
                w.WriteAttribute("s", styleId);

                if (dataType is not null)
                    w.WriteAttributeString("t", dataType);

                if (xlCell.ShowPhonetic)
                    w.WriteAttributeString("ph", TrueValue);

                if (xlCell.CellMetaIndex is not null)
                    w.WriteAttribute("cm", xlCell.CellMetaIndex.Value);

                if (xlCell.ValueMetaIndex is not null)
                    w.WriteAttribute("vm", xlCell.ValueMetaIndex.Value);
            }

            static void WriteCell(XmlWriter xml, XlsxCell xlCell, char[] cellRef, SaveContext context, SaveOptions options, HashSet<IXlsxAddress> tableTotalCells, uint rowStyleId)
            {
                var styleId = context.GetStyleId(xlCell.GetFormat());

                Span<Char> cellRefSpan = cellRef;
                var cellRefLen = xlCell.Point.Format(cellRefSpan);

                if (xlCell.HasFormula)
                {
                    String dataType = null;
                    if (options.EvaluateFormulasBeforeSaving)
                    {
                        try
                        {
                            xlCell.Evaluate(false);
                            dataType = FormulaDataType[(int)xlCell.DataType];
                        }
                        catch
                        {
                            // Do nothing, cell will be left blank. Unimplemented features or functions would stop trying to save a file.
                        }
                    }

                    WriteStartCell(xml, xlCell, cellRef, cellRefLen, dataType, styleId);

                    var xlFormula = xlCell.Formula;
                    if (xlFormula.Type == FormulaType.DataTable)
                    {
                        // Data table doesn't write actual text of formula, that is referenced by context
                        xml.WriteStartElement("f", Main2006SsNs);
                        xml.WriteAttributeString("t", "dataTable");
                        xml.WriteAttributeString("ref", xlFormula.Range.ToString());

                        var is2D = xlFormula.Is2DDataTable;
                        if (is2D)
                            xml.WriteAttributeString("dt2D", TrueValue);

                        var isDataRowTable = xlFormula.IsRowDataTable;
                        if (isDataRowTable)
                            xml.WriteAttributeString("dtr", TrueValue);

                        xml.WriteAttributeString("r1", xlFormula.Input1.ToString());
                        var input1Deleted = xlFormula.Input1Deleted;
                        if (input1Deleted)
                            xml.WriteAttributeString("del1", TrueValue);

                        if (is2D)
                            xml.WriteAttributeString("r2", xlFormula.Input2.ToString());

                        var input2Deleted = xlFormula.Input2Deleted;
                        if (input2Deleted)
                            xml.WriteAttributeString("del2", TrueValue);

                        // Excel doesn't recalculate table formula on load or on click of a button or any kind of forced recalculation.
                        // It is necessary to mark some precedent formula dirty (e.g. edit cell formula and enter in Excel).
                        // By setting the CalculateCell, we ensure that Excel will calculate values of data table formula on load and
                        // user will see correct values.
                        xml.WriteAttributeString("ca", TrueValue);

                        xml.WriteEndElement(); // f
                    }
                    else if (xlCell.HasArrayFormula)
                    {
                        var isMasterCell = xlCell.Formula.Range.FirstPoint == xlCell.Point;
                        if (isMasterCell)
                        {
                            xml.WriteStartElement("f", Main2006SsNs);
                            xml.WriteAttributeString("t", "array");
                            xml.WriteAttributeString("ref", xlCell.FormulaReference.ToStringRelative());
                            xml.WriteString(xlCell.FormulaA1);
                            xml.WriteEndElement(); // f
                        }
                    }
                    else
                    {
                        xml.WriteStartElement("f", Main2006SsNs);
                        xml.WriteString(xlCell.FormulaA1);
                        xml.WriteEndElement(); // f
                    }

                    if (options.EvaluateFormulasBeforeSaving && xlCell.CachedValue.Type != XlsxDataType.Blank && !xlCell.NeedsRecalculation)
                    {
                        WriteCellValue(xml, xlCell, context);
                    }

                    xml.WriteEndElement(); // cell
                }
                else if (tableTotalCells.Contains(xlCell.Address))
                {
                    var table = xlCell.Worksheet.Tables.First<XlsxTable>(t => t.AsRange().Contains(xlCell));
                    var field = (XlsxTableField)table.Fields.First(f => f.Column.ColumnNumber() == xlCell.Address.ColumnNumber);

                    // If this is a cell in the totals row that contains a label (xor with function), write label
                    // Only label can be written. Total functions are basically formulas that use structured
                    // references and SR are not yet supported, so not yet possible to calculate total values.
                    if (!String.IsNullOrWhiteSpace(field.TotalsRowLabel))
                    {
                        // Excel requires that table totals row label attribute in tableColumn must match the cell
                        // string from SST. If they don't match, Excel will consider it a corrupt workbook.
                        var sharedStringId = context.GetSharedStringId(xlCell, field.TotalsRowLabel);
                        WriteStartCell(xml, xlCell, cellRef, cellRefLen, "s", styleId);
                        xml.WriteStartElement("v", Main2006SsNs);
                        xml.WriteValue(sharedStringId);
                        xml.WriteEndElement();
                    }
                    xml.WriteEndElement(); // cell
                }
                else if (xlCell.DataType != XlsxDataType.Blank)
                {
                    // Cell contains only a value
                    var dataType = GetCellValueType(xlCell);
                    WriteStartCell(xml, xlCell, cellRef, cellRefLen, dataType, styleId);

                    WriteCellValue(xml, xlCell, context);
                    xml.WriteEndElement(); // cell
                }
                else if (rowStyleId != styleId)
                {
                    // Cell is blank and should be written only if it has different style from parent.
                    // Non-written cells use inherited style of a row.
                    WriteStartCell(xml, xlCell, cellRef, cellRefLen, null, styleId);
                    xml.WriteEndElement(); // cell
                }
            }
        }

        /// <summary>
        /// An array to convert data type for a formula cell. Key is <see cref="XlsxDataType"/>.
        /// It saves some performance through direct indexation instead of switch.
        /// </summary>
        private static readonly String[] FormulaDataType =
        {
            null, // blank
            "b", // boolean
            null, // number, default value, no need to save type
            "str", // text, formula can only save this type, no inline or shared string
            "e", // error
            null, // datetime, saved as serialized date-time
            null // timespan, saved as serialized date-time
        };

        /// <summary>
        /// An array to convert data type for a cell that only contains a value. Key is <see cref="XlsxDataType"/>.
        /// It saves some performance through direct indexation instead of switch.
        /// </summary>
        private static readonly String[] ValueDataType =
        {
            null, // blank
            "b", // boolean
            null, // number, default value, no need to save type
            "s", // text, the default is a shared string, but there also can be inline string depending on ShareString property
            "e", // error
            null, // datetime, saved as serialized date-time
            null // timespan, saved as serialized date-time
        };

        private static String GetCellValueType(XlsxCell xlCell)
        {
            var dataType = xlCell.DataType;
            if (dataType == XlsxDataType.Text && !xlCell.ShareString)
                return "inlineStr";
            return ValueDataType[(int)dataType];
        }

        private static Int32 GetMaxColumn(XlsxWorksheet xlWorksheet)
        {
            var maxColumn = 0;

            if (!xlWorksheet.Internals.CellsCollection.IsEmpty)
            {
                maxColumn = xlWorksheet.Internals.CellsCollection.MaxColumnUsed;
            }

            if (xlWorksheet.Internals.ColumnsCollection.Count > 0)
            {
                var maxColCollection = xlWorksheet.Internals.ColumnsCollection.Keys.Max();
                if (maxColCollection > maxColumn)
                    maxColumn = maxColCollection;
            }

            return maxColumn;
        }
    }
}
