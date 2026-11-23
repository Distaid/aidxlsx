#nullable disable

using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PageSetup;

internal class XlsxPageSetup : IXlsxPageSetup
{
    public XlsxPageSetup(XlsxPageSetup defaultPageOptions, XlsxWorksheet worksheet)
    {

        if (defaultPageOptions != null)
        {
            PrintAreas = new XlsxPrintAreas(defaultPageOptions.PrintAreas as XlsxPrintAreas, worksheet);
            CenterHorizontally = defaultPageOptions.CenterHorizontally;
            CenterVertically = defaultPageOptions.CenterVertically;
            FirstPageNumber = defaultPageOptions.FirstPageNumber;
            HorizontalDpi = defaultPageOptions.HorizontalDpi;
            PageOrientation = defaultPageOptions.PageOrientation;
            VerticalDpi = defaultPageOptions.VerticalDpi;
            FirstRowToRepeatAtTop = defaultPageOptions.FirstRowToRepeatAtTop;
            LastRowToRepeatAtTop = defaultPageOptions.LastRowToRepeatAtTop;
            FirstColumnToRepeatAtLeft = defaultPageOptions.FirstColumnToRepeatAtLeft;
            LastColumnToRepeatAtLeft = defaultPageOptions.LastColumnToRepeatAtLeft;
            ShowComments = defaultPageOptions.ShowComments;


            PaperSize = defaultPageOptions.PaperSize;
            _pagesTall = defaultPageOptions.PagesTall;
            _pagesWide = defaultPageOptions.PagesWide;
            _scale = defaultPageOptions.Scale;


            if (defaultPageOptions.Margins != null)
            {
                Margins = new XlsxMargins
                {
                    Top = defaultPageOptions.Margins.Top,
                    Bottom = defaultPageOptions.Margins.Bottom,
                    Left = defaultPageOptions.Margins.Left,
                    Right = defaultPageOptions.Margins.Right,
                    Header = defaultPageOptions.Margins.Header,
                    Footer = defaultPageOptions.Margins.Footer
                };
            }
            AlignHFWithMargins = defaultPageOptions.AlignHFWithMargins;
            ScaleHFWithDocument = defaultPageOptions.ScaleHFWithDocument;
            ShowGridlines = defaultPageOptions.ShowGridlines;
            ShowRowAndColumnHeadings = defaultPageOptions.ShowRowAndColumnHeadings;
            BlackAndWhite = defaultPageOptions.BlackAndWhite;
            DraftQuality = defaultPageOptions.DraftQuality;
            PageOrder = defaultPageOptions.PageOrder;

            ColumnBreaks = [.. defaultPageOptions.ColumnBreaks];
            RowBreaks = [.. defaultPageOptions.RowBreaks];
            Header = new XlsxHeaderFooter(defaultPageOptions.Header as XlsxHeaderFooter, worksheet);
            Footer = new XlsxHeaderFooter(defaultPageOptions.Footer as XlsxHeaderFooter, worksheet);
            PrintErrorValue = defaultPageOptions.PrintErrorValue;
        }
        else
        {
            PrintAreas = new XlsxPrintAreas(worksheet);
            Header = new XlsxHeaderFooter(worksheet);
            Footer = new XlsxHeaderFooter(worksheet);
            ColumnBreaks = [];
            RowBreaks = [];
        }
    }
    public IXlsxPrintAreas PrintAreas { get; private set; }


    public int FirstRowToRepeatAtTop { get; private set; }
    public int LastRowToRepeatAtTop { get; private set; }
    public void SetRowsToRepeatAtTop(string range)
    {
        var arrRange = range.Replace("$", "").Split(':');
        SetRowsToRepeatAtTop(int.Parse(arrRange[0]), int.Parse(arrRange[1]));
    }
    public void SetRowsToRepeatAtTop(int firstRowToRepeatAtTop, int lastRowToRepeatAtTop)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(firstRowToRepeatAtTop);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(firstRowToRepeatAtTop, lastRowToRepeatAtTop);

        FirstRowToRepeatAtTop = firstRowToRepeatAtTop;
        LastRowToRepeatAtTop = lastRowToRepeatAtTop;
    }
    public int FirstColumnToRepeatAtLeft { get; private set; }
    public int LastColumnToRepeatAtLeft { get; private set; }
    public void SetColumnsToRepeatAtLeft(string range)
    {
        var arrRange = range.Replace("$", "").Split(':');
        if (int.TryParse(arrRange[0], out _))
            SetColumnsToRepeatAtLeft(int.Parse(arrRange[0]), int.Parse(arrRange[1]));
        else
            SetColumnsToRepeatAtLeft(arrRange[0], arrRange[1]);
    }
    public void SetColumnsToRepeatAtLeft(string firstColumnToRepeatAtLeft, string lastColumnToRepeatAtLeft)
    {
        SetColumnsToRepeatAtLeft(XlsxHelper.GetColumnNumberFromLetter(firstColumnToRepeatAtLeft), XlsxHelper.GetColumnNumberFromLetter(lastColumnToRepeatAtLeft));
    }
    public void SetColumnsToRepeatAtLeft(int firstColumnToRepeatAtLeft, int lastColumnToRepeatAtLeft)
    {
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(firstColumnToRepeatAtLeft);
        ArgumentOutOfRangeException.ThrowIfGreaterThan(firstColumnToRepeatAtLeft, lastColumnToRepeatAtLeft);

        FirstColumnToRepeatAtLeft = firstColumnToRepeatAtLeft;
        LastColumnToRepeatAtLeft = lastColumnToRepeatAtLeft;
    }

    public XlsxPageOrientation PageOrientation { get; set; }
    public XlsxPaperSize PaperSize { get; set; }
    public int HorizontalDpi { get; set; }
    public int VerticalDpi { get; set; }
    public int? FirstPageNumber { get; set; }
    public bool CenterHorizontally { get; set; }
    public bool CenterVertically { get; set; }
    public XlsxPrintErrorValues PrintErrorValue { get; set; }
    public IXlsxMargins Margins { get; set; }

    private int _pagesWide;
    public int PagesWide
    {
        get
        {
            return _pagesWide;
        }
        set
        {
            _pagesWide = value;
            if (_pagesWide > 0)
                _scale = 0;
        }
    }

    private int _pagesTall;
    public int PagesTall
    {
        get
        {
            return _pagesTall;
        }
        set
        {
            _pagesTall = value;
            if (_pagesTall > 0)
                _scale = 0;
        }
    }

    private int _scale;
    public int Scale
    {
        get
        {
            return _scale;
        }
        set
        {
            _scale = value;
            if (_scale <= 0) return;
            _pagesTall = 0;
            _pagesWide = 0;
        }
    }

    public void AdjustTo(int percentageOfNormalSize)
    {
        Scale = percentageOfNormalSize;
        _pagesWide = 0;
        _pagesTall = 0;
    }
    public void FitToPages(int pagesWide, int pagesTall)
    {
        _pagesWide = pagesWide;
        this._pagesTall = pagesTall;
        _scale = 0;
    }


    public IXlsxHeaderFooter Header { get; private set; }
    public IXlsxHeaderFooter Footer { get; private set; }

    public bool ScaleHFWithDocument { get; set; }
    public bool AlignHFWithMargins { get; set; }

    public bool ShowGridlines { get; set; }
    public bool ShowRowAndColumnHeadings { get; set; }
    public bool BlackAndWhite { get; set; }
    public bool DraftQuality { get; set; }

    public XlsxPageOrderValues PageOrder { get; set; }
    public XlsxShowCommentsValues ShowComments { get; set; }

    public List<int> RowBreaks { get; private set; }
    public List<int> ColumnBreaks { get; private set; }
    public void AddHorizontalPageBreak(int row)
    {
        if (!RowBreaks.Contains(row))
            RowBreaks.Add(row);
        RowBreaks.Sort();
    }
    public void AddVerticalPageBreak(int column)
    {
        if (!ColumnBreaks.Contains(column))
            ColumnBreaks.Add(column);
        ColumnBreaks.Sort();
    }

    //public void SetPageBreak(IXlsxRange range, XlsxPageBreakLocations breakLocation)
    //{
    //    switch (breakLocation)
    //    {
    //        case XlsxPageBreakLocations.AboveRange: RowBreaks.Add(range.Internals.Worksheet.Row(range.RowNumber)); break;
    //        case XlsxPageBreakLocations.BelowRange: RowBreaks.Add(range.Internals.Worksheet.Row(range.RowCount())); break;
    //        case XlsxPageBreakLocations.LeftOfRange: ColumnBreaks.Add(range.Internals.Worksheet.Column(range.ColumnNumber)); break;
    //        case XlsxPageBreakLocations.RightOfRange: ColumnBreaks.Add(range.Internals.Worksheet.Column(range.ColumnCount())); break;
    //        default: throw new NotImplementedException();
    //    }
    //}

    public IXlsxPageSetup SetPageOrientation(XlsxPageOrientation value) { PageOrientation = value; return this; }
    public IXlsxPageSetup SetPagesWide(int value) { PagesWide = value; return this; }
    public IXlsxPageSetup SetPagesTall(int value) { PagesTall = value; return this; }
    public IXlsxPageSetup SetScale(int value) { Scale = value; return this; }
    public IXlsxPageSetup SetHorizontalDpi(int value) { HorizontalDpi = value; return this; }
    public IXlsxPageSetup SetVerticalDpi(int value) { VerticalDpi = value; return this; }
    public IXlsxPageSetup SetFirstPageNumber(int? value) { FirstPageNumber = value; return this; }
    public IXlsxPageSetup SetCenterHorizontally() { CenterHorizontally = true; return this; }
    public IXlsxPageSetup SetCenterHorizontally(bool value) { CenterHorizontally = value; return this; }
    public IXlsxPageSetup SetCenterVertically() { CenterVertically = true; return this; }
    public IXlsxPageSetup SetCenterVertically(bool value) { CenterVertically = value; return this; }
    public IXlsxPageSetup SetPaperSize(XlsxPaperSize value) { PaperSize = value; return this; }
    public IXlsxPageSetup SetScaleHFWithDocument() { ScaleHFWithDocument = true; return this; }
    public IXlsxPageSetup SetScaleHFWithDocument(bool value) { ScaleHFWithDocument = value; return this; }
    public IXlsxPageSetup SetAlignHFWithMargins() { AlignHFWithMargins = true; return this; }
    public IXlsxPageSetup SetAlignHFWithMargins(bool value) { AlignHFWithMargins = value; return this; }
    public IXlsxPageSetup SetShowGridlines() { ShowGridlines = true; return this; }
    public IXlsxPageSetup SetShowGridlines(bool value) { ShowGridlines = value; return this; }
    public IXlsxPageSetup SetShowRowAndColumnHeadings() { ShowRowAndColumnHeadings = true; return this; }
    public IXlsxPageSetup SetShowRowAndColumnHeadings(bool value) { ShowRowAndColumnHeadings = value; return this; }
    public IXlsxPageSetup SetBlackAndWhite() { BlackAndWhite = true; return this; }
    public IXlsxPageSetup SetBlackAndWhite(bool value) { BlackAndWhite = value; return this; }
    public IXlsxPageSetup SetDraftQuality() { DraftQuality = true; return this; }
    public IXlsxPageSetup SetDraftQuality(bool value) { DraftQuality = value; return this; }
    public IXlsxPageSetup SetPageOrder(XlsxPageOrderValues value) { PageOrder = value; return this; }
    public IXlsxPageSetup SetShowComments(XlsxShowCommentsValues value) { ShowComments = value; return this; }
    public IXlsxPageSetup SetPrintErrorValue(XlsxPrintErrorValues value) { PrintErrorValue = value; return this; }

    public bool DifferentFirstPageOnHF { get; set; }
    public IXlsxPageSetup SetDifferentFirstPageOnHF()
    {
        return SetDifferentFirstPageOnHF(true);
    }
    public IXlsxPageSetup SetDifferentFirstPageOnHF(bool value)
    {
        DifferentFirstPageOnHF = value;
        return this;
    }
    public bool DifferentOddEvenPagesOnHF { get; set; }
    public IXlsxPageSetup SetDifferentOddEvenPagesOnHF()
    {
        return SetDifferentOddEvenPagesOnHF(true);
    }
    public IXlsxPageSetup SetDifferentOddEvenPagesOnHF(bool value)
    {
        DifferentOddEvenPagesOnHF = value;
        return this;
    }
}
