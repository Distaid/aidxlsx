#nullable disable

using DocumentFormat.OpenXml.Spreadsheet;

namespace Aid.Xlsx.Internal.IO
{
    internal enum XlsxSheetViewContents
    {
        Pane,
        Selection,
        PivotSelection,
        ExtensionList
    }

    internal class XlsxSheetViewContentManager : XlsxBaseContentManager<XlsxSheetViewContents>
    {
        public XlsxSheetViewContentManager(SheetView sheetView)
        {
            contents.Add(XlsxSheetViewContents.Pane, sheetView.Elements<Pane>().LastOrDefault());
            contents.Add(XlsxSheetViewContents.Selection, sheetView.Elements<Selection>().LastOrDefault());
            contents.Add(XlsxSheetViewContents.PivotSelection, sheetView.Elements<PivotSelection>().LastOrDefault());
            contents.Add(XlsxSheetViewContents.ExtensionList, sheetView.Elements<ExtensionList>().LastOrDefault());
        }
    }
}
