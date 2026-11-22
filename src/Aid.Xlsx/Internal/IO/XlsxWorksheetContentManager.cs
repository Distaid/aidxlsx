#nullable disable

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using SpreadsheetColumns = DocumentFormat.OpenXml.Spreadsheet.Columns;
using SpreadsheetCustomProperties = DocumentFormat.OpenXml.Spreadsheet.CustomProperties;
using SpreadsheetHyperlinks = DocumentFormat.OpenXml.Spreadsheet.Hyperlinks;
using SpreadsheetPageSetup = DocumentFormat.OpenXml.Spreadsheet.PageSetup;

namespace Aid.Xlsx.Internal.IO
{
    internal enum XlsxWorksheetContents
    {
        SheetProperties = 1,
        SheetDimension = 2,
        SheetViews = 3,
        SheetFormatProperties = 4,
        Columns = 5,
        SheetData = 6,
        SheetCalculationProperties = 7,
        SheetProtection = 8,
        ProtectedRanges = 9,
        Scenarios = 10,
        AutoFilter = 11,
        SortState = 12,
        DataConsolidate = 13,
        CustomSheetViews = 14,
        MergeCells = 15,
        PhoneticProperties = 16,
        ConditionalFormatting = 17,
        DataValidations = 18,
        Hyperlinks = 19,
        PrintOptions = 20,
        PageMargins = 21,
        PageSetup = 22,
        HeaderFooter = 23,
        RowBreaks = 24,
        ColumnBreaks = 25,
        CustomProperties = 26,
        CellWatches = 27,
        IgnoredErrors = 28,
        SmartTags = 29,
        Drawing = 30,
        LegacyDrawing = 31,
        LegacyDrawingHeaderFooter = 32,
        DrawingHeaderFooter = 33,
        Picture = 34,
        OleObjects = 35,
        Controls = 36,
        AlternateContent = 37,
        WebPublishItems = 38,
        TableParts = 39,
        WorksheetExtensionList = 40
    }

    internal class XlsxWorksheetContentManager : XlsxBaseContentManager<XlsxWorksheetContents>
    {
        public XlsxWorksheetContentManager(Worksheet opWorksheet)
        {
            contents.Add(XlsxWorksheetContents.SheetProperties, opWorksheet.Elements<SheetProperties>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.SheetDimension, opWorksheet.Elements<SheetDimension>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.SheetViews, opWorksheet.Elements<SheetViews>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.SheetFormatProperties, opWorksheet.Elements<SheetFormatProperties>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.Columns, opWorksheet.Elements<SpreadsheetColumns>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.SheetData, opWorksheet.Elements<SheetData>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.SheetCalculationProperties, opWorksheet.Elements<SheetCalculationProperties>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.SheetProtection, opWorksheet.Elements<SheetProtection>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.ProtectedRanges, opWorksheet.Elements<ProtectedRanges>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.Scenarios, opWorksheet.Elements<Scenarios>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.AutoFilter, opWorksheet.Elements<AutoFilter>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.SortState, opWorksheet.Elements<SortState>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.DataConsolidate, opWorksheet.Elements<DataConsolidate>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.CustomSheetViews, opWorksheet.Elements<CustomSheetViews>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.MergeCells, opWorksheet.Elements<MergeCells>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.PhoneticProperties, opWorksheet.Elements<PhoneticProperties>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.ConditionalFormatting, opWorksheet.Elements<ConditionalFormatting>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.DataValidations, opWorksheet.Elements<DataValidations>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.Hyperlinks, opWorksheet.Elements<SpreadsheetHyperlinks>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.PrintOptions, opWorksheet.Elements<PrintOptions>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.PageMargins, opWorksheet.Elements<PageMargins>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.PageSetup, opWorksheet.Elements<SpreadsheetPageSetup>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.HeaderFooter, opWorksheet.Elements<HeaderFooter>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.RowBreaks, opWorksheet.Elements<RowBreaks>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.ColumnBreaks, opWorksheet.Elements<ColumnBreaks>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.CustomProperties, opWorksheet.Elements<SpreadsheetCustomProperties>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.CellWatches, opWorksheet.Elements<CellWatches>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.IgnoredErrors, opWorksheet.Elements<IgnoredErrors>().LastOrDefault());
            //contents.Add(XlsxWSContents.SmartTags, opWorksheet.Elements<SmartTags>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.Drawing, opWorksheet.Elements<Drawing>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.LegacyDrawing, opWorksheet.Elements<LegacyDrawing>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.LegacyDrawingHeaderFooter, opWorksheet.Elements<LegacyDrawingHeaderFooter>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.DrawingHeaderFooter, opWorksheet.Elements<DrawingHeaderFooter>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.Picture, opWorksheet.Elements<Picture>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.OleObjects, opWorksheet.Elements<OleObjects>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.Controls, opWorksheet.Elements<Controls>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.AlternateContent, opWorksheet.Elements<AlternateContent>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.WebPublishItems, opWorksheet.Elements<WebPublishItems>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.TableParts, opWorksheet.Elements<TableParts>().LastOrDefault());
            contents.Add(XlsxWorksheetContents.WorksheetExtensionList, opWorksheet.Elements<WorksheetExtensionList>().LastOrDefault());
        }
    }
}
