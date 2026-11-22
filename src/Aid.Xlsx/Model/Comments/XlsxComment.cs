#nullable disable warnings

using System.Diagnostics;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.Drawings.Styles;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Rows;
using Aid.Xlsx.Model.RichText;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.Comments;

internal class XlsxComment : XlsxFormattedText<IXlsxComment>, IXlsxComment
{
    private XlsxCell _cell;

    private static XlsxFontFormatValue DefaultCommentFont
    {
        get
        {
            // MS Excel uses Tahoma 9 Swiss no matter what current style font
            var defaultCommentFont = XlsxFontFormatValue.Default with
            {
                Name = "Tahoma",
                Size = XlsxFontSize.FromPoints(9),
                Family = XlsxFontFamilyNumberingValues.Swiss,
                Color = XlsxColor.Black
            };

            return defaultCommentFont;
        }
    }

    private XlsxComment(XlsxFontFormatValue defaultFont, XlsxWorkbookStyles styles, XlsxFontFormatValue? phoneticsFont)
        : base(defaultFont, styles)
    {
        if (phoneticsFont is not null)
        {
            Debug.Assert(styles.Fonts.ContainsValue(phoneticsFont));
            Phonetics = new XlsxPhonetics(phoneticsFont, defaultFont, styles, OnContentChanged);
        }
    }

    #region IXlsxComment Members

    public String Author { get; set; }

    public IXlsxComment SetAuthor(String value)
    {
        Author = value;
        return this;
    }

    public IXlsxRichString AddSignature()
    {
        AddText(Author + ":").SetBold();
        return AddText(Environment.NewLine);
    }

    public void Delete()
    {
        _cell.DeleteComment();
    }

    #endregion IXlsxComment Members

    #region IXlsxDrawing

    public String Name { get; set; }
    public String Description { get; set; }
    public XlsxDrawingAnchor Anchor { get; set; }
    public Boolean HorizontalFlip { get; set; }
    public Boolean VerticalFlip { get; set; }
    public Int32 Rotation { get; set; }
    public Int32 ExtentLength { get; set; }
    public Int32 ExtentWidth { get; set; }
    public Int32 ShapeId { get; internal set; }
    public Boolean Visible { get; set; }

    public IXlsxComment SetVisible()
    {
        Visible = true;
        return Container;
    }

    public IXlsxComment SetVisible(Boolean hidden)
    {
        Visible = hidden;
        return Container;
    }

    public IXlsxDrawingPosition Position { get; private set; }

    public Int32 ZOrder { get; set; }

    public IXlsxComment SetZOrder(Int32 zOrder)
    {
        ZOrder = zOrder;
        return Container;
    }

    public IXlsxDrawingStyle Style { get; private set; }

    public IXlsxComment SetName(String name)
    {
        Name = name;
        return Container;
    }

    public IXlsxComment SetDescription(String description)
    {
        Description = description;
        return Container;
    }

    public IXlsxComment SetHorizontalFlip()
    {
        HorizontalFlip = true;
        return Container;
    }

    public IXlsxComment SetHorizontalFlip(Boolean horizontalFlip)
    {
        HorizontalFlip = horizontalFlip;
        return Container;
    }

    public IXlsxComment SetVerticalFlip()
    {
        VerticalFlip = true;
        return Container;
    }

    public IXlsxComment SetVerticalFlip(Boolean verticalFlip)
    {
        VerticalFlip = verticalFlip;
        return Container;
    }

    public IXlsxComment SetRotation(Int32 rotation)
    {
        Rotation = rotation;
        return Container;
    }

    public IXlsxComment SetExtentLength(Int32 extentLength)
    {
        ExtentLength = extentLength;
        return Container;
    }

    public IXlsxComment SetExtentWidth(Int32 extentWidth)
    {
        ExtentWidth = extentWidth;
        return Container;
    }

    #endregion IXlsxDrawing

    internal static XlsxComment Create(XlsxCell cell, int? shapeId)
    {
        var styles = cell.Worksheet.Workbook.Styles;
        var defaultFont = styles.RegisterFontFormat(DefaultCommentFont);
        var comment = new XlsxComment(defaultFont, styles, null);
        comment.Initialize(cell, shapeId: shapeId);
        return comment;
    }

    internal static XlsxComment CreateAsCopy(XlsxCell targetCell, XlsxCell sourceCell, XlsxComment originalComment)
    {
        // source cell could be from different workbook, so register formats
        var styles = targetCell.Worksheet.Workbook.Styles;
        var defaultFont = styles.RegisterFontFormat(sourceCell.GetFormat().Font);
        var phoneticsFont = originalComment.HasPhonetics
            ? styles.RegisterFontFormat(XlsxFontFormatValue.FromFontBase(originalComment.Phonetics, styles))
            : null;
        var comment = new XlsxComment(defaultFont, styles, phoneticsFont);

        foreach (XlsxRichString rt in originalComment)
            comment.AddText(rt.Text, rt);

        comment.Initialize(targetCell, originalComment.Style);
        return comment;
    }

    private void Initialize(XlsxCell cell, IXlsxDrawingStyle style = null, int? shapeId = null)
    {
        style = style ?? XlsxDrawingStyle.DefaultCommentStyle;
        shapeId = shapeId ?? cell.Worksheet.Workbook.ShapeIdManager.GetNext();

        Author = cell.Worksheet.Author;
        Container = this;
        Anchor = XlsxDrawingAnchor.MoveAndSizeWithCells;
        Style = new XlsxDrawingStyle();
        Int32 previousRowNumber = cell.Address.RowNumber;
        Double previousRowOffset = 0;

        if (previousRowNumber > 1)
        {
            previousRowNumber--;

            if (cell.Worksheet.Internals.RowsCollection.TryGetValue(previousRowNumber, out XlsxRow previousRow))
                previousRowOffset = Math.Max(0, previousRow.Height - 7);
            else
                previousRowOffset = Math.Max(0, cell.Worksheet.RowHeight - 7);
        }

        Position = new XlsxDrawingPosition
        {
            Column = cell.Address.ColumnNumber + 1,
            ColumnOffset = 2,
            Row = previousRowNumber,
            RowOffset = previousRowOffset
        };

        ZOrder = cell.Worksheet.ZOrder++;
        Style
            .Margins.SetLeft(style.Margins.Left)
            .Margins.SetRight(style.Margins.Right)
            .Margins.SetTop(style.Margins.Top)
            .Margins.SetBottom(style.Margins.Bottom)
            .Margins.SetAutomatic(style.Margins.Automatic)
            .Size.SetHeight(style.Size.Height)
            .Size.SetWidth(style.Size.Width)
            .ColorsAndLines.SetLineColor(style.ColorsAndLines.LineColor)
            .ColorsAndLines.SetFillColor(style.ColorsAndLines.FillColor)
            .ColorsAndLines.SetLineDash(style.ColorsAndLines.LineDash)
            .ColorsAndLines.SetLineStyle(style.ColorsAndLines.LineStyle)
            .ColorsAndLines.SetLineWeight(style.ColorsAndLines.LineWeight)
            .ColorsAndLines.SetFillTransparency(style.ColorsAndLines.FillTransparency)
            .ColorsAndLines.SetLineTransparency(style.ColorsAndLines.LineTransparency)
            .Alignment.SetHorizontal(style.Alignment.Horizontal)
            .Alignment.SetVertical(style.Alignment.Vertical)
            .Alignment.SetDirection(style.Alignment.Direction)
            .Alignment.SetOrientation(style.Alignment.Orientation)
            .Alignment.SetAutomaticSize(style.Alignment.AutomaticSize)
            .Properties.SetPositioning(style.Properties.Positioning)
            .Protection.SetLocked(style.Protection.Locked)
            .Protection.SetLockText(style.Protection.LockText);

        _cell = cell;
        ShapeId = shapeId.Value;
    }
}
