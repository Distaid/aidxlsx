#nullable disable

using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingStyle : IXlsxDrawingStyle
{
    public XlsxDrawingStyle()
    {
        //Font = new XlsxDrawingFont(this);
        Alignment = new XlsxDrawingAlignment(this);
        ColorsAndLines = new XlsxDrawingColorsAndLines(this);
        Size = new XlsxDrawingSize(this);
        Protection = new XlsxDrawingProtection(this);
        Properties = new XlsxDrawingProperties(this);
        Margins = new XlsxDrawingMargins(this);
        Web = new XlsxDrawingWeb(this);
    }

    //public IXlsxDrawingFont Font { get; private set; }
    public IXlsxDrawingAlignment Alignment { get; private set; }

    public IXlsxDrawingColorsAndLines ColorsAndLines { get; private set; }
    public IXlsxDrawingSize Size { get; private set; }
    public IXlsxDrawingProtection Protection { get; private set; }
    public IXlsxDrawingProperties Properties { get; private set; }
    public IXlsxDrawingMargins Margins { get; private set; }
    public IXlsxDrawingWeb Web { get; private set; }

    public static IXlsxDrawingStyle DefaultCommentStyle
    {
        get
        {
            var defaultCommentStyle = new XlsxDrawingStyle();

            defaultCommentStyle
                .Margins.SetLeft(0.1)
                .Margins.SetRight(0.1)
                .Margins.SetTop(0.05)
                .Margins.SetBottom(0.05)
                .Margins.SetAutomatic()
                .Size.SetHeight(59.25)
                .Size.SetWidth(19.2)
                .ColorsAndLines.SetLineColor(XlsxColor.Black)
                .ColorsAndLines.SetFillColor(XlsxColor.FromArgb(255, 255, 225))
                .ColorsAndLines.SetLineDash(XlsxDashStyle.Solid)
                .ColorsAndLines.SetLineStyle(XlsxLineStyle.Single)
                .ColorsAndLines.SetLineWeight(0.75)
                .ColorsAndLines.SetFillTransparency(1)
                .ColorsAndLines.SetLineTransparency(1)
                .Alignment.SetHorizontal(XlsxDrawingHorizontalAlignment.Left)
                .Alignment.SetVertical(XlsxDrawingVerticalAlignment.Top)
                .Alignment.SetDirection(XlsxDrawingTextDirection.LeftToRight)
                .Alignment.SetOrientation(XlsxDrawingTextOrientation.LeftToRight)
                .Properties.SetPositioning(XlsxDrawingAnchor.Absolute)
                .Protection.SetLocked()
                .Protection.SetLockText();

            return defaultCommentStyle;
        }
    }
}
