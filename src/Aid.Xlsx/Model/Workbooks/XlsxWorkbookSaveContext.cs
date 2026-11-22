#nullable disable

using System.Diagnostics;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Tables;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Opc.Packaging;
using Aid.Xlsx.SpreadsheetML.Workbooks;

namespace Aid.Xlsx.Model.Workbooks;

public partial class XlsxWorkbook
{
    #region Nested type: SaveContext

    internal sealed class SaveContext
    {
        public SaveContext()
        {
            RelIdGenerator = new RelIdGenerator();
            TableId = 0;
            TableNames = new HashSet<String>();
        }

        public RelIdGenerator RelIdGenerator { get; }

        /// <summary>
        /// A map of number format to a number format id for saved file. It contains all number
        /// formats from the file, all number formats used in the application (styles, pivot
        /// tables, dxf) and all predefined formats.
        /// </summary>
        internal Dictionary<string, int> NumberFormatMap = new();

        internal Dictionary<XlsxFontFormatValue, int> FontMap = new();

        internal Dictionary<XlsxCellFormatValue, uint> FormatMap = new();

        internal Dictionary<XlsxDxfValue, uint> DxfMap = new();

        public uint TableId { get; set; }
        public HashSet<string> TableNames { get; private set; }

        /// <summary>
        /// A free id that can be used by the workbook to reference to a pivot cache.
        /// The <c>PivotCaches</c> element in a workbook connects the parts with pivot
        /// cache parts.
        /// </summary>

        /// <summary>
        /// A map of shared string ids. The index is the actual index from sharedStringId and
        /// value is an mapped stringId to write to a file. The mapped stringId has no gaps
        /// between ids.
        /// </summary>
        public List<int> SstMap { get; set; }

#nullable enable
        internal int GetSharedStringId(XlsxCell xlCell, string text)
        {
            var sharedStringId = SstMap[xlCell.MemorySstId];
            if (sharedStringId < 0)
            {
                throw new UnreachableException($"Unable to find text '{text}' in shared string table for cell {xlCell.Point}. " +
                                               "That likely means reference counting is broken. As a stop-gap, try to set the " +
                                               "text value to an unused cell to increase number of references for the text.");
            }

            return sharedStringId;
        }

        /// <summary>
        /// Get id of number format that is going to be actually saved to database.
        /// </summary>
        internal int? GetNumberFormatId(string? numberFormat)
        {
            if (numberFormat is null)
                return null;

            return NumberFormatMap[numberFormat];
        }

        internal int GetFontId(XlsxFontFormatValue font)
        {
            return FontMap[font];
        }

        internal uint GetDxfId(XlsxDxfValue dxf)
        {
            return DxfMap[dxf];
        }

        internal uint GetStyleId(XlsxCellFormatValue? format)
        {
            return format is not null ? FormatMap[format] : 0;
        }
#nullable disable
    }

    #endregion Nested type: SaveContext

    #region Nested type: RelType

    internal enum RelType
    {
        Workbook//, Worksheet
    }

    #endregion Nested type: RelType

    #region Nested type: RelIdGenerator

    internal sealed class RelIdGenerator
    {
        private readonly Dictionary<RelType, HashSet<String>> _relIds = new();

        public void AddValues(IEnumerable<String> values, RelType relType)
        {
            if (!_relIds.TryGetValue(relType, out var set))
            {
                set = new HashSet<string>();
                _relIds.Add(relType, set);
            }

            set.UnionWith(values);
        }

        /// <summary>
        /// Add all existing rel ids present on the parts or workbook to the generator, so they are not duplicated again.
        /// </summary>
        public void AddExistingValues(WorkbookPackage package, OpcPart workbookPart, XlsxWorkbook xlWorkbook)
        {
            AddValues(workbookPart.Relationships.Select(rel => rel.Id), RelType.Workbook);
            AddValues(xlWorkbook.WorksheetsInternal.Cast<XlsxWorksheet>().Where(ws => !String.IsNullOrWhiteSpace(ws.RelId)).Select(ws => ws.RelId), RelType.Workbook);
            AddValues(xlWorkbook.WorksheetsInternal.Cast<XlsxWorksheet>().Where(ws => !String.IsNullOrWhiteSpace(ws.LegacyDrawingId)).Select(ws => ws.LegacyDrawingId), RelType.Workbook);
            AddValues(xlWorkbook.WorksheetsInternal
                .Cast<XlsxWorksheet>()
                .SelectMany(ws => ws.Tables.Cast<XlsxTable>())
                .Where(t => !String.IsNullOrWhiteSpace(t.RelId))
                .Select(t => t.RelId), RelType.Workbook);

            foreach (var xlWorksheet in xlWorkbook.WorksheetsInternal.Cast<XlsxWorksheet>())
            {
                // if the worksheet is a new one, it doesn't have RelId yet.
                if (string.IsNullOrEmpty(xlWorksheet.RelId) || package.RelatedPart(workbookPart, xlWorksheet.RelId, "worksheet") is not { } worksheetPart)
                    continue;

                AddValues(worksheetPart.Relationships.Select(rel => rel.Id), RelType.Workbook);
                if (package.FirstRelatedPart(worksheetPart, "drawing").Part is { } drawingPart)
                    AddValues(drawingPart.Relationships.Select(rel => rel.Id), RelType.Workbook);
            }
        }

        public String GetNext(RelType relType)
        {
            if (!_relIds.TryGetValue(relType, out var set))
            {
                set = new HashSet<String>();
                _relIds.Add(relType, set);
            }

            var id = set.Count + 1;
            while (true)
            {
                var relId = String.Concat("rId", id);
                if (!set.Contains(relId))
                {
                    set.Add(relId);
                    return relId;
                }
                id++;
            }
        }
    }

    #endregion Nested type: RelIdGenerator
}
