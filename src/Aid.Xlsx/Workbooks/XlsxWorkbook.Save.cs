#nullable disable

using System.Diagnostics;
using System.Globalization;
using System.Xml.Linq;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.Internal.IO;
using Aid.Xlsx.Opc;
using Aid.Xlsx.PivotTables;
using Aid.Xlsx.Worksheets;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;
using Path = System.IO.Path;

namespace Aid.Xlsx.Workbooks
{
    public partial class XlsxWorkbook
    {
        private void Validate(SpreadsheetDocument package)
        {
            var backupCulture = Thread.CurrentThread.CurrentCulture;

            IList<ValidationErrorInfo> errors;
            try
            {
                Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
                var validator = new OpenXmlValidator();
                errors = validator.Validate(package).ToArray();

                // The styles part is written directly into the part stream, without intermediate OOXML SDK
                // representation. The validation loads the written XML into a memory so it can validate it.
                // But the loaded in-memory represenation from validation is then picked up by
                // the SpreadsheetDocument saving code and the XML would be re-serialized and the original XML
                // from part writer would be discarded. That is a problem for the following reasons:
                // * XML of a part would be different when saved with a validation and without a validation
                // * There is no control over XML serialization. The writer is attempting to mimic Excel for
                //   easier comparison against Excel (e.g. default namespace doesn't have a prefix), but part
                //   writer just uses default serialization setting.
                // To solve this, we discard the in-memory represenation created by the validator. Thus the save
                // code will not re-serialize the part that has already been written.
                if (package.WorkbookPart.WorkbookStylesPart is { } stylesPart)
                    stylesPart.UnloadRootElement();
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = backupCulture;
            }

            if (errors.Any())
            {
                var message = string.Join("\r\n", errors.Select(e => string.Format("Part {0}, Path {1}: {2}", e.Part.Uri, e.Path.XPath, e.Description)).ToArray());
                throw new ApplicationException(message);
            }
        }

        private void CreatePackage(String filePath, SpreadsheetDocumentType spreadsheetDocumentType, SaveOptions options)
        {
            var directoryName = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directoryName)) Directory.CreateDirectory(directoryName);

            var exists = File.Exists(filePath);
            using var stream = File.Open(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            WritePackage(stream, !exists, exists ? spreadsheetDocumentType : null, spreadsheetDocumentType, options);
        }

        private void CreatePackage(Stream stream, bool newStream, SpreadsheetDocumentType spreadsheetDocumentType, SaveOptions options)
        {
            WritePackage(stream, newStream, null, spreadsheetDocumentType, options);
        }

        private void WritePackage(Stream stream, bool newStream, SpreadsheetDocumentType? requestedType,
            SpreadsheetDocumentType spreadsheetDocumentType, SaveOptions options)
        {
            using var opc = newStream ? OpcPackage.Create() : OpcPackage.Open(stream);
            using var package = new OpcSdkPackage(opc, FileAccess.ReadWrite);
            string newExtendedRelationshipId;
            using (var document = newStream
                ? SpreadsheetDocument.Create(package, spreadsheetDocumentType)
                : SpreadsheetDocument.Open(package))
            {
                if (requestedType is { } type && document.DocumentType != type)
                    document.ChangeDocumentType(type);

                newExtendedRelationshipId = CreateParts(document, options, WorkbookPropertiesParts.HasExtendedProperties(opc));
                if (options.ValidatePackage) Validate(document);
            }

            package.SynchronizeParts();
            WorkbookPropertiesParts.Write(opc, this, newExtendedRelationshipId);

            if (newStream)
            {
                opc.Save(stream);
                return;
            }

            using var output = new MemoryStream();
            opc.Save(output);
            stream.Position = 0;
            stream.SetLength(0);
            output.Position = 0;
            output.CopyTo(stream);
        }

        private static void DeleteSheetAndDependencies(WorkbookPart workbookPart, WorksheetPart worksheetPart, string sheetName)
        {
            if (sheetName is not null)
            {
                var cachesFromDeletedSheet = workbookPart.PivotTableCacheDefinitionParts
                    .Where(part => part.PivotCacheDefinition.CacheSource?.WorksheetSource?.Sheet == sheetName)
                    .ToList();
                foreach (var cache in cachesFromDeletedSheet)
                    workbookPart.DeletePart(cache);
            }

            workbookPart.DeletePart(worksheetPart);
        }

        // Adds child parts and generates content of the specified part.
        private string CreateParts(SpreadsheetDocument document, SaveOptions options, bool hasExtendedProperties)
        {
            var context = new SaveContext();

            var workbookPart = document.WorkbookPart ?? document.AddWorkbookPart();
            XElement workbookRoot;
            using (var existing = workbookPart.GetStream(FileMode.OpenOrCreate, FileAccess.Read))
                workbookRoot = existing.Length == 0 ? WorkbookPartWriter.CreateRoot() : WorkbookXml.ReadRoot(existing);

            var worksheets = WorksheetsInternal;
            var partsToRemove = workbookPart.Parts.Where(s => worksheets.Deleted.Contains(s.RelationshipId)).ToList();
            var originalSheets = workbookRoot.Element(workbookRoot.Name.Namespace + "sheets")?.Elements(workbookRoot.Name.Namespace + "sheet")
                .Select(WorkbookXml.ReadSheet).Where(sheet => !string.IsNullOrEmpty(sheet.Id)).ToDictionary(sheet => sheet.Id) ?? [];
            foreach (var part in partsToRemove)
            {
                if (part.OpenXmlPart is WorksheetPart worksheetPart)
                    DeleteSheetAndDependencies(workbookPart, worksheetPart, originalSheets.GetValueOrDefault(part.RelationshipId).Name);
            }

            // Ensure all RelId's have been added to the context
            context.RelIdGenerator.AddExistingValues(workbookPart, this);

            // Keep existing relationship numbering while the workbook and root parts share an ID generator.
            var newExtendedRelationshipId = hasExtendedProperties ? null : context.RelIdGenerator.GetNext(RelType.Workbook);

            WorkbookPartWriter.GenerateContent(workbookRoot, this, options, context);

            var workbookStylesPart = workbookPart.WorkbookStylesPart ??
                                     workbookPart.AddNewPart<WorkbookStylesPart>(
                                         context.RelIdGenerator.GetNext(RelType.Workbook));

            using (var target = workbookStylesPart.GetStream(FileMode.Create, FileAccess.Write))
                new StylesWriter().WriteContent(target, XmlToEnumMapper.Instance, Styles, this, context);

            var sharedStringTablePart = workbookPart.SharedStringTablePart ??
                                        workbookPart.AddNewPart<SharedStringTablePart>(
                                            context.RelIdGenerator.GetNext(RelType.Workbook));

            using (var target = sharedStringTablePart.GetStream(FileMode.Create, FileAccess.Write))
                SharedStringTableWriter.GenerateSharedStringTablePartContent(this, target, context);

            var cacheRelIds = PivotCachesInternal
                  .Select<XlsxPivotCache, String>(ps => ps.WorkbookCacheRelId)
                  .Where(relId => !string.IsNullOrWhiteSpace(relId))
                  .Distinct();

            foreach (var relId in cacheRelIds)
            {
                if (workbookPart.TryGetPartById(relId, out var part) && part is PivotTableCacheDefinitionPart pivotTableCacheDefinitionPart)
                    pivotTableCacheDefinitionPart.PivotCacheDefinition.CacheFields.RemoveAllChildren();
            }

            var allPivotTables = WorksheetsInternal.SelectMany<XlsxWorksheet, IXlsxPivotTable>(ws => ws.PivotTables).ToList();

            // Phase 1 - Synchronize all pivot cache parts in the document, so each
            // source that will be saved has all required parts created and relationship
            // ids are set (in this case `Workbook.PivotCaches` relationship table).
            // Only sources that are used by a table are saved.
            SynchronizePivotTableParts(workbookPart, allPivotTables, context);
            WorkbookPartWriter.WritePivotCaches(workbookRoot, allPivotTables.Select(table => table.PivotCache).Distinct().Cast<XlsxPivotCache>().ToList());
            using (var target = workbookPart.GetStream(FileMode.Create, FileAccess.Write))
                new XDocument(workbookRoot).Save(target, System.Xml.Linq.SaveOptions.DisableFormatting);

            // Phase 2 - All parts and relationships are set, fill in the parts.
            if (allPivotTables.Any())
            {
                GeneratePivotCaches(workbookPart, context);
            }

            foreach (var worksheet in WorksheetsInternal.Cast<XlsxWorksheet>().OrderBy(w => w.Position))
            {
                WorksheetPart worksheetPart;
                var wsRelId = worksheet.RelId;
                bool partIsEmpty;
                if (workbookPart.Parts.Any(p => p.RelationshipId == wsRelId))
                {
                    worksheetPart = (WorksheetPart)workbookPart.GetPartById(wsRelId);
                    partIsEmpty = false;
                }
                else
                {
                    worksheetPart = workbookPart.AddNewPart<WorksheetPart>(wsRelId);
                    partIsEmpty = true;
                }

                var worksheetHasComments = worksheet.Internals.CellsCollection.GetCells(c => c.HasComment).Any();

                var commentsPart = worksheetPart.WorksheetCommentsPart;

                // VML part is the source of truth for shapes of comments, form controls and likely others.
                // Excel won't display any shape without VML. The drawing part is always present, but is likely
                // only different rendering of VML (more precisely the shapes behind VML).
                var vmlDrawingPart = worksheetPart.VmlDrawingParts.FirstOrDefault();
                XDocument vmlContent;
                using (var original = vmlDrawingPart?.GetStream(FileMode.Open, FileAccess.Read))
                    vmlContent = VmlDrawingPartWriter.GenerateContent(original, worksheet, worksheetHasComments);

                if (worksheetHasComments)
                {
                    // If sheet has comments, we must keep VML in legacy drawing part to display them
                    // as well as comments part for semantic reasons.
                    if (commentsPart == null)
                    {
                        commentsPart = worksheetPart.AddNewPart<WorksheetCommentsPart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                    }

                    if (vmlDrawingPart == null)
                    {
                        if (String.IsNullOrWhiteSpace(worksheet.LegacyDrawingId))
                        {
                            worksheet.LegacyDrawingId = context.RelIdGenerator.GetNext(RelType.Workbook);
                        }

                        vmlDrawingPart = worksheetPart.AddNewPart<VmlDrawingPart>(worksheet.LegacyDrawingId);
                    }

                    using (var target = commentsPart.GetStream(FileMode.Create, FileAccess.Write))
                        CommentPartWriter.GenerateWorksheetCommentsPartContent(target, worksheet);
                }
                else
                {
                    // There are no comments in the worksheet = the comment part is no longer needed,
                    // but VML part might contain other shapes, like form controls.
                    if (commentsPart is not null)
                        worksheetPart.DeletePart(commentsPart);
                }

                if (vmlContent is not null && vmlDrawingPart is not null)
                {
                    using var target = vmlDrawingPart.GetStream(FileMode.Create, FileAccess.Write);
                    VmlDrawingPartWriter.Write(target, vmlContent);
                }
                else if (vmlDrawingPart is not null)
                {
                    worksheet.LegacyDrawingId = null;
                    worksheetPart.DeletePart(vmlDrawingPart);
                }

                var xlTables = worksheet.Tables;

                // The way forward is to have 2-phase save, this is a start of that
                // concept for tables:
                //
                // Phase 1 - synchronize part existence with tables xlWorksheet, so each
                // table has a corresponding part and part that don't are deleted.
                // This phase doesn't modify the content, it only ensures that RelIds are set
                // corresponding parts exist and the parts that don't exist are removed
                TablePartWriter.SynchronizeTableParts(xlTables, worksheetPart, context);

                // Phase 2 - At this point, all pieces must have corresponding parts
                // The only way to link between parts is through RelIds that were already
                // set in phase 1. The phase 2 is all about content of individual parts.
                // Each part should have individual writer.
                TablePartWriter.GenerateTableParts(xlTables, worksheetPart, context);

                WorksheetPartWriter.GenerateWorksheetPartContent(partIsEmpty, worksheetPart, worksheet, options, context);

                if (worksheet.PivotTables.Any<XlsxPivotTable>())
                {
                    GeneratePivotTables(workbookPart, worksheetPart, worksheet, context);
                }
            }

            var calculationChainPart = workbookPart.CalculationChainPart;
            XDocument calculationChain = null;
            if (options.GenerateCalculationChain)
            {
                using var existing = calculationChainPart?.GetStream(FileMode.Open, FileAccess.Read);
                calculationChain = CalculationChainPartWriter.GenerateContent(existing, this);
            }

            if (calculationChain is not null)
            {
                calculationChainPart ??= workbookPart.AddNewPart<CalculationChainPart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                using var target = calculationChainPart.GetStream(FileMode.Create, FileAccess.Write);
                calculationChain.Save(target, System.Xml.Linq.SaveOptions.DisableFormatting);
            }
            else if (calculationChainPart is not null)
            {
                workbookPart.DeletePart(calculationChainPart);
            }

            if (workbookPart.ThemePart == null)
            {
                var themePart = workbookPart.AddNewPart<ThemePart>(context.RelIdGenerator.GetNext(RelType.Workbook));
                using var target = themePart.GetStream(FileMode.Create, FileAccess.Write);
                ThemeXml.WriteDefault(target, Theme);
            }

            // Clear list of deleted worksheets to prevent errors on multiple saves
            worksheets.Deleted.Clear();
            return newExtendedRelationshipId;
        }

        private static void SynchronizePivotTableParts(WorkbookPart workbookPart, IReadOnlyList<IXlsxPivotTable> allPivotTables, SaveContext context)
        {
            RemoveUnusedPivotCacheDefinitionParts(workbookPart, allPivotTables);
            AddUsedPivotCacheDefinitionParts(workbookPart, allPivotTables, context);

            // Remove pivot cache parts that are a part of the loaded document, but aren't used by a pivot table of the xlWorkbook
            // part of the first phase of saving
            static void RemoveUnusedPivotCacheDefinitionParts(WorkbookPart workbookPart, IReadOnlyList<IXlsxPivotTable> allPivotTables)
            {
                var workbookCacheRelIds = allPivotTables
                    .Select(pt => pt.PivotCache.CastTo<XlsxPivotCache>().WorkbookCacheRelId)
                    .Distinct()
                    .ToList();

                var orphanedParts = workbookPart
                    .GetPartsOfType<PivotTableCacheDefinitionPart>()
                    .Where(pcdp => !workbookCacheRelIds.Contains(workbookPart.GetIdOfPart(pcdp)))
                    .ToList();

                foreach (var orphanPart in orphanedParts)
                {
                    orphanPart.DeletePart(orphanPart.PivotTableCacheRecordsPart);
                    workbookPart.DeletePart(orphanPart);
                }

            }

            static void AddUsedPivotCacheDefinitionParts(WorkbookPart workbookPart, IReadOnlyList<IXlsxPivotTable> allPivotTables, SaveContext context)
            {
                // Add ids and part for the caches to workbooks
                // We might get a XlsxPivotSource with an id of apart that isn't in the file (e.g. loaded from a file and saved to a different one).
                var newPivotSources = allPivotTables
                    .Select(pt => pt.PivotCache.CastTo<XlsxPivotCache>())
                    .Where(ps => string.IsNullOrEmpty(ps.WorkbookCacheRelId) || !workbookPart.HasPartWithId(ps.WorkbookCacheRelId))
                    .Distinct()
                    .ToList();

                foreach (var pivotSource in newPivotSources)
                {
                    var cacheRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    pivotSource.WorkbookCacheRelId = cacheRelId;

                    workbookPart.AddNewPart<PivotTableCacheDefinitionPart>(pivotSource.WorkbookCacheRelId);
                }
            }
        }

        private void GeneratePivotCaches(WorkbookPart workbookPart, SaveContext context)
        {
            var pivotTables = WorksheetsInternal.SelectMany<XlsxWorksheet, XlsxPivotTable>(ws => ws.PivotTables);

            var xlPivotCaches = pivotTables.Select(pt => pt.PivotCache).Distinct();
            foreach (var xlPivotCache in xlPivotCaches)
            {
                Debug.Assert(!string.IsNullOrEmpty(xlPivotCache.WorkbookCacheRelId));

                var pivotTableCacheDefinitionPart = (PivotTableCacheDefinitionPart)workbookPart.GetPartById(xlPivotCache.WorkbookCacheRelId);

                PivotTableCacheDefinitionPartWriter.GenerateContent(pivotTableCacheDefinitionPart, xlPivotCache, context);

                var pivotTableCacheRecordsPart = pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Any()
                    ? pivotTableCacheDefinitionPart.GetPartsOfType<PivotTableCacheRecordsPart>().Single()
                    : pivotTableCacheDefinitionPart.AddNewPart<PivotTableCacheRecordsPart>("rId1");

                using var target = pivotTableCacheRecordsPart.GetStream(FileMode.Create, FileAccess.Write);
                PivotCacheRecordsWriter.WriteContent(target, xlPivotCache);
            }
        }

        private static void GeneratePivotTables(
            WorkbookPart workbookPart,
            WorksheetPart worksheetPart,
            XlsxWorksheet xlWorksheet,
            SaveContext context)
        {
            foreach (var pt in xlWorksheet.PivotTables)
            {
                PivotTablePart pivotTablePart;
                var createNewPivotTablePart = String.IsNullOrWhiteSpace(pt.RelId);
                if (createNewPivotTablePart)
                {
                    var relId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    pt.RelId = relId;
                    pivotTablePart = worksheetPart.AddNewPart<PivotTablePart>(relId);
                }
                else
                    pivotTablePart = (PivotTablePart)worksheetPart.GetPartById(pt.RelId);

                var pivotSource = pt.PivotCache;
                var pivotTableCacheDefinitionPart = pivotTablePart.PivotTableCacheDefinitionPart;
                if (!workbookPart.GetPartById(pivotSource.WorkbookCacheRelId).Equals(pivotTableCacheDefinitionPart))
                {
                    pivotTablePart.DeletePart(pivotTableCacheDefinitionPart);
                    pivotTablePart.CreateRelationshipToPart(workbookPart.GetPartById(pivotSource.WorkbookCacheRelId), context.RelIdGenerator.GetNext(XlsxWorkbook.RelType.Workbook));
                }

                PivotTableDefinitionPartWriter2.WriteContent(pivotTablePart, pt, context);
            }
        }

    }
}
