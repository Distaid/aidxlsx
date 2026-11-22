#nullable disable

using System.Text;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.RichText;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;

namespace Aid.Xlsx.Model.PageSetup;

internal class XlsxHFText
{
    private readonly XlsxHFItem _hfItem;

    public XlsxHFText(XlsxRichString richText, XlsxHFItem hfItem)
    {
        RichText = richText;
        _hfItem = hfItem;
    }

    public XlsxRichString RichText { get; private set; }

    public String GetHFText(String prevText)
    {
        var wsFont = _hfItem.HeaderFooter.Worksheet.Style.Font;

        var isRichText = RichText.FontName != null && RichText.FontName != wsFont.FontName
                         || RichText.Bold != wsFont.Bold
                         || RichText.Italic != wsFont.Italic
                         || RichText.Strikethrough != wsFont.Strikethrough
                         || RichText.FontSize > 0 && Math.Abs(RichText.FontSize - wsFont.FontSize) > XlsxHelper.Epsilon
                         || RichText.VerticalAlignment != wsFont.VerticalAlignment
                         || RichText.Underline != wsFont.Underline
                         || !RichText.FontColor.Equals(wsFont.FontColor);

        if (!isRichText)
            return RichText.Text;

        StringBuilder sb = new StringBuilder();

        if (RichText.FontName != null && RichText.FontName != wsFont.FontName)
            sb.Append("&\"" + RichText.FontName);
        else
            sb.Append("&\"-");

        if (RichText.Bold && RichText.Italic)
            sb.Append(",Bold Italic\"");
        else if (RichText.Bold)
            sb.Append(",Bold\"");
        else if (RichText.Italic)
            sb.Append(",Italic\"");
        else
            sb.Append(",Regular\"");

        if (RichText.FontSize > 0 && Math.Abs(RichText.FontSize - wsFont.FontSize) > XlsxHelper.Epsilon)
            sb.Append("&" + RichText.FontSize);

        if (RichText.Strikethrough && !wsFont.Strikethrough)
            sb.Append("&S");

        if (RichText.VerticalAlignment != wsFont.VerticalAlignment)
        {
            if (RichText.VerticalAlignment == XlsxFontVerticalTextAlignmentValues.Subscript)
                sb.Append("&Y");
            else if (RichText.VerticalAlignment == XlsxFontVerticalTextAlignmentValues.Superscript)
                sb.Append("&X");
        }

        if (RichText.Underline != wsFont.Underline)
        {
            if (RichText.Underline == XlsxFontUnderlineValues.Single)
                sb.Append("&U");
            else if (RichText.Underline == XlsxFontUnderlineValues.Double)
                sb.Append("&E");
        }

        var lastColorPosition = prevText.LastIndexOf("&K");

        if (
            (lastColorPosition >= 0 && !RichText.FontColor.Equals(XlsxColor.FromHtml("#" + prevText.Substring(lastColorPosition + 2, 6))))
            || (lastColorPosition == -1 && !RichText.FontColor.Equals(wsFont.FontColor))
            )
            sb.Append("&K" + RichText.FontColor.Color.ToHex().Substring(2));

        sb.Append(RichText.Text);

        if (RichText.Underline != wsFont.Underline)
        {
            if (RichText.Underline == XlsxFontUnderlineValues.Single)
                sb.Append("&U");
            else if (RichText.Underline == XlsxFontUnderlineValues.Double)
                sb.Append("&E");
        }

        if (RichText.VerticalAlignment != wsFont.VerticalAlignment)
        {
            if (RichText.VerticalAlignment == XlsxFontVerticalTextAlignmentValues.Subscript)
                sb.Append("&Y");
            else if (RichText.VerticalAlignment == XlsxFontVerticalTextAlignmentValues.Superscript)
                sb.Append("&X");
        }

        if (RichText.Strikethrough && !wsFont.Strikethrough)
            sb.Append("&S");

        return sb.ToString();
    }
}
