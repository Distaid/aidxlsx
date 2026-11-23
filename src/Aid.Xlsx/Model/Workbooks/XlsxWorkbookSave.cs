#nullable disable

using System.Diagnostics;
using System.Xml.Linq;
using Path = System.IO.Path;
using Aid.Xlsx.Model.PivotTables;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Opc.Packaging;
using Aid.Xlsx.SpreadsheetML.Drawings;
using Aid.Xlsx.SpreadsheetML.Pivots;
using Aid.Xlsx.SpreadsheetML.Properties;
using Aid.Xlsx.SpreadsheetML.Strings;
using Aid.Xlsx.SpreadsheetML.Styles;
using Aid.Xlsx.SpreadsheetML.Tables;
using Aid.Xlsx.SpreadsheetML.Workbooks;
using Aid.Xlsx.SpreadsheetML.Worksheets;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.Model.Workbooks;

public partial class XlsxWorkbook
{
    private void CreatePackage(String filePath, WorkbookPackageType spreadsheetDocumentType, SaveOptions options)
    {
        var directoryName = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(directoryName)) Directory.CreateDirectory(directoryName);

        var exists = File.Exists(filePath);
        using var stream = File.Open(filePath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
        WritePackage(stream, !exists, exists ? spreadsheetDocumentType : null, spreadsheetDocumentType, options);
    }

    private void CreatePackage(Stream stream, bool newStream, WorkbookPackageType spreadsheetDocumentType, SaveOptions options)
    {
        WritePackage(stream, newStream, null, spreadsheetDocumentType, options);
    }

    private void WritePackage(Stream stream, bool newStream, WorkbookPackageType? requestedType,
        WorkbookPackageType spreadsheetDocumentType, SaveOptions options)
    {
        using var opc = newStream ? OpcPackage.Create() : OpcPackage.Open(stream);
        var package = new WorkbookPackage(opc);
        var workbookPart = newStream
            ? package.AddWorkbook(GetWorkbookContentType(spreadsheetDocumentType))
            : package.WorkbookPart ?? throw new InvalidDataException("The package has no workbook part.");
        if (requestedType is { } type)
            workbookPart.ChangeContentType(GetWorkbookContentType(type));

        var newExtendedRelationshipId = CreateParts(package, workbookPart, options, WorkbookPropertiesParts.HasExtendedProperties(opc));
        WorkbookPropertiesParts.Write(opc, this, newExtendedRelationshipId);
        if (options.ValidatePackage)
            package.Validate(workbookPart);

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

    private static string GetWorkbookContentType(WorkbookPackageType type) => type switch
    {
        WorkbookPackageType.Workbook => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        WorkbookPackageType.Template => "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        WorkbookPackageType.MacroEnabledWorkbook => "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        WorkbookPackageType.MacroEnabledTemplate => "application/vnd.ms-excel.template.macroEnabled.main+xml",
        _ => throw new ArgumentOutOfRangeException(nameof(type)),
    };

    private static void DeleteSheetAndDependencies(WorkbookPackage package, OpcPart workbookPart, string id, string sheetName)
    {
        if (sheetName is not null)
        {
            var cachesFromDeletedSheet = package.RelatedParts(workbookPart, "pivotCacheDefinition")
                .Where(pair => ReferencesSheet(pair.Part, sheetName))
                .ToList();
            foreach (var (cacheId, _) in cachesFromDeletedSheet)
                package.Unlink(workbookPart, cacheId);
        }

        package.Unlink(workbookPart, id);

        static bool ReferencesSheet(OpcPart part, string name)
        {
            using var stream = part.OpenRead();
            var definition = PivotCacheDefinitionXml.Read(stream);
            var ns = definition.Name.Namespace;
            return (string)definition.Element(ns + "cacheSource")?.Element(ns + "worksheetSource")?.Attribute("sheet") == name;
        }
    }

    // Adds child parts and generates content of the specified part.
    private string CreateParts(WorkbookPackage package, OpcPart workbookPart, SaveOptions options, bool hasExtendedProperties)
    {
        var context = new SaveContext();

        XElement workbookRoot;
        using (var existing = workbookPart.OpenRead())
            workbookRoot = workbookPart.Length == 0 ? WorkbookPartWriter.CreateRoot() : WorkbookXml.ReadRoot(existing);

        var worksheets = WorksheetsInternal;
        var partsToRemove = package.RelatedParts(workbookPart, "worksheet").Where(pair => worksheets.Deleted.Contains(pair.Id)).ToList();
        var originalSheets = workbookRoot.Element(workbookRoot.Name.Namespace + "sheets")?.Elements(workbookRoot.Name.Namespace + "sheet")
            .Select(WorkbookXml.ReadSheet).Where(sheet => !string.IsNullOrEmpty(sheet.Id)).ToDictionary(sheet => sheet.Id) ?? [];
        foreach (var (id, _) in partsToRemove)
            DeleteSheetAndDependencies(package, workbookPart, id, originalSheets.GetValueOrDefault(id).Name);

        // Ensure all RelId's have been added to the context
        context.RelIdGenerator.AddExistingValues(package, workbookPart, this);

        // Keep existing relationship numbering while the workbook and root parts share an ID generator.
        var newExtendedRelationshipId = hasExtendedProperties ? null : context.RelIdGenerator.GetNext(RelType.Workbook);

        WorkbookPartWriter.GenerateContent(workbookRoot, this, options, context);

        var workbookStylesPart = package.FirstRelatedPart(workbookPart, "styles").Part ??
                                 package.AddPart(workbookPart, context.RelIdGenerator.GetNext(RelType.Workbook), "styles");

        using (var target = workbookStylesPart.OpenWrite())
            new StylesWriter().WriteContent(target, XmlToEnumMapper.Instance, Styles, this, context);

        var sharedStringTablePart = package.FirstRelatedPart(workbookPart, "sharedStrings").Part ??
                                    package.AddPart(workbookPart, context.RelIdGenerator.GetNext(RelType.Workbook), "sharedStrings");

        using (var target = sharedStringTablePart.OpenWrite())
            SharedStringTableWriter.GenerateSharedStringTablePartContent(this, target, context);

        var allPivotTables = WorksheetsInternal.SelectMany<XlsxWorksheet, IXlsxPivotTable>(ws => ws.PivotTables).ToList();

        // Phase 1 - Synchronize all pivot cache parts in the document, so each
        // source that will be saved has all required parts created and relationship
        // ids are set (in this case `Workbook.PivotCaches` relationship table).
        // Only sources that are used by a table are saved.
        SynchronizePivotTableParts(package, workbookPart, allPivotTables, context);
        WorkbookPartWriter.WritePivotCaches(workbookRoot, [.. allPivotTables.Select(table => table.PivotCache).Distinct().Cast<XlsxPivotCache>()]);
        using (var target = workbookPart.OpenWrite())
            new XDocument(workbookRoot).Save(target, System.Xml.Linq.SaveOptions.DisableFormatting);

        // Phase 2 - All parts and relationships are set, fill in the parts.
        if (allPivotTables.Count != 0)
        {
            GeneratePivotCaches(package, workbookPart);
        }

        foreach (var worksheet in WorksheetsInternal.Cast<XlsxWorksheet>().OrderBy(w => w.Position))
        {
            var wsRelId = worksheet.RelId;
            var worksheetPart = package.RelatedPart(workbookPart, wsRelId, "worksheet") ?? package.AddPart(workbookPart, wsRelId, "worksheet");

            var worksheetHasComments = worksheet.Internals.CellsCollection.GetCells(c => c.HasComment).Any();

            var (commentsId, commentsPart) = package.FirstRelatedPart(worksheetPart, "comments");

            // VML part is the source of truth for shapes of comments, form controls and likely others.
            // Excel won't display any shape without VML. The drawing part is always present, but is likely
            // only different rendering of VML (more precisely the shapes behind VML).
            var (vmlId, vmlDrawingPart) = package.FirstRelatedPart(worksheetPart, "vmlDrawing");
            XDocument vmlContent;
            using (var original = vmlDrawingPart?.OpenRead())
                vmlContent = VmlDrawingPartWriter.GenerateContent(original, worksheet, worksheetHasComments);

            if (worksheetHasComments)
            {
                // If sheet has comments, we must keep VML in legacy drawing part to display them
                // as well as comments part for semantic reasons.
                commentsPart ??= package.AddPart(worksheetPart, context.RelIdGenerator.GetNext(RelType.Workbook), "comments");

                if (vmlDrawingPart == null)
                {
                    if (String.IsNullOrWhiteSpace(worksheet.LegacyDrawingId))
                    {
                        worksheet.LegacyDrawingId = context.RelIdGenerator.GetNext(RelType.Workbook);
                    }

                    vmlDrawingPart = package.AddPart(worksheetPart, worksheet.LegacyDrawingId, "vmlDrawing");
                }

                using var target = commentsPart.OpenWrite();
                CommentPartWriter.GenerateWorksheetCommentsPartContent(target, worksheet);
            }
            else
            {
                // There are no comments in the worksheet = the comment part is no longer needed,
                // but VML part might contain other shapes, like form controls.
                if (commentsPart is not null)
                    package.Unlink(worksheetPart, commentsId);
            }

            if (vmlContent is not null && vmlDrawingPart is not null)
            {
                using var target = vmlDrawingPart.OpenWrite();
                VmlDrawingPartWriter.Write(target, vmlContent);
            }
            else if (vmlDrawingPart is not null)
            {
                worksheet.LegacyDrawingId = null;
                package.Unlink(worksheetPart, vmlId);
            }

            var xlTables = worksheet.Tables;

            // The way forward is to have 2-phase save, this is a start of that
            // concept for tables:
            //
            // Phase 1 - synchronize part existence with tables xlWorksheet, so each
            // table has a corresponding part and part that don't are deleted.
            // This phase doesn't modify the content, it only ensures that RelIds are set
            // corresponding parts exist and the parts that don't exist are removed
            TablePartWriter.SynchronizeTableParts(xlTables, package, worksheetPart, context);

            // Phase 2 - At this point, all pieces must have corresponding parts
            // The only way to link between parts is through RelIds that were already
            // set in phase 1. The phase 2 is all about content of individual parts.
            // Each part should have individual writer.
            TablePartWriter.GenerateTableParts(xlTables, package, worksheetPart, context);

            WorksheetPartWriter.GenerateWorksheetPartContent(package, worksheetPart, worksheet, options, context);

            if (worksheet.PivotTables.Any<XlsxPivotTable>())
            {
                GeneratePivotTables(package, workbookPart, worksheetPart, worksheet, context);
            }
        }

        var (calculationChainId, calculationChainPart) = package.FirstRelatedPart(workbookPart, "calcChain");
        XDocument calculationChain = null;
        if (options.GenerateCalculationChain)
        {
            using var existing = calculationChainPart?.OpenRead();
            calculationChain = CalculationChainPartWriter.GenerateContent(existing, this);
        }

        if (calculationChain is not null)
        {
            calculationChainPart ??= package.AddPart(workbookPart, context.RelIdGenerator.GetNext(RelType.Workbook), "calcChain");
            using var target = calculationChainPart.OpenWrite();
            calculationChain.Save(target, System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        else if (calculationChainPart is not null)
        {
            package.Unlink(workbookPart, calculationChainId);
        }

        if (package.FirstRelatedPart(workbookPart, "theme").Part is null)
        {
            var themePart = package.AddPart(workbookPart, context.RelIdGenerator.GetNext(RelType.Workbook), "theme");
            using var target = themePart.OpenWrite();
            ThemeXml.WriteDefault(target, Theme);
        }

        // Clear list of deleted worksheets to prevent errors on multiple saves
        worksheets.Deleted.Clear();
        return newExtendedRelationshipId;
    }

    private static void SynchronizePivotTableParts(WorkbookPackage package, OpcPart workbookPart,
        IReadOnlyList<IXlsxPivotTable> allPivotTables, SaveContext context)
    {
        var caches = allPivotTables.Select(table => (XlsxPivotCache)table.PivotCache).Distinct().ToList();
        var usedIds = caches.Select(cache => cache.WorkbookCacheRelId).ToHashSet(StringComparer.Ordinal);
        foreach (var (id, _) in package.RelatedParts(workbookPart, "pivotCacheDefinition").ToList())
            if (!usedIds.Contains(id))
                package.Unlink(workbookPart, id);

        foreach (var cache in caches)
        {
            if (!string.IsNullOrEmpty(cache.WorkbookCacheRelId) &&
                package.RelatedPart(workbookPart, cache.WorkbookCacheRelId, "pivotCacheDefinition") is not null)
                continue;

            cache.WorkbookCacheRelId = context.RelIdGenerator.GetNext(RelType.Workbook);
            package.AddPart(workbookPart, cache.WorkbookCacheRelId, "pivotCacheDefinition");
        }
    }

    private void GeneratePivotCaches(WorkbookPackage package, OpcPart workbookPart)
    {
        var pivotTables = WorksheetsInternal.SelectMany<XlsxWorksheet, XlsxPivotTable>(ws => ws.PivotTables);

        var xlPivotCaches = pivotTables.Select(pt => pt.PivotCache).Distinct();
        foreach (var xlPivotCache in xlPivotCaches)
        {
            Debug.Assert(!string.IsNullOrEmpty(xlPivotCache.WorkbookCacheRelId));

            var definitionPart = package.RelatedPart(workbookPart, xlPivotCache.WorkbookCacheRelId, "pivotCacheDefinition");

            XElement definition;
            using (var existing = definitionPart.OpenRead())
                definition = definitionPart.Length == 0 ? null : PivotCacheDefinitionXml.Read(existing);

            var (recordsId, recordsPart) = package.FirstRelatedPart(definitionPart, "pivotCacheRecords");
            if (recordsPart is null)
            {
                recordsId = GetRecordsRelationshipId(definitionPart);
                recordsPart = package.AddPart(definitionPart, recordsId, "pivotCacheRecords");
            }
            using (var definitionTarget = definitionPart.OpenWrite())
                PivotCacheDefinitionXml.Write(definitionTarget, definition, xlPivotCache, recordsId);

            using var target = recordsPart.OpenWrite();
            PivotCacheRecordsWriter.WriteContent(target, xlPivotCache);
        }

        static string GetRecordsRelationshipId(OpcPart part)
        {
            for (var index = 1; ; ++index)
            {
                var id = FormattableString.Invariant($"rId{index}");
                if (!part.Relationships.TryGet(id, out _))
                    return id;
            }
        }
    }

    private static void GeneratePivotTables(
        WorkbookPackage package,
        OpcPart workbookPart,
        OpcPart worksheetPart,
        XlsxWorksheet xlWorksheet,
        SaveContext context)
    {
        foreach (var pt in xlWorksheet.PivotTables)
        {
            var pivotTablePart = string.IsNullOrWhiteSpace(pt.RelId) ? null : package.RelatedPart(worksheetPart, pt.RelId, "pivotTable");
            if (pivotTablePart is null)
            {
                pt.RelId = context.RelIdGenerator.GetNext(RelType.Workbook);
                pivotTablePart = package.AddPart(worksheetPart, pt.RelId, "pivotTable");
            }

            var workbookCache = package.RelatedPart(workbookPart, pt.PivotCache.WorkbookCacheRelId, "pivotCacheDefinition");
            var (cacheId, pivotTableCache) = package.FirstRelatedPart(pivotTablePart, "pivotCacheDefinition");
            if (workbookCache != pivotTableCache)
            {
                if (cacheId is not null)
                    package.Unlink(pivotTablePart, cacheId);
                do
                    cacheId = context.RelIdGenerator.GetNext(RelType.Workbook);
                while (pivotTablePart.Relationships.TryGet(cacheId, out _));
                WorkbookPackage.Link(pivotTablePart, cacheId, "pivotCacheDefinition", workbookCache);
            }

            using var stream = pivotTablePart.OpenWrite();
            PivotTableXml.Write(stream, pt, context);
        }
    }

}
