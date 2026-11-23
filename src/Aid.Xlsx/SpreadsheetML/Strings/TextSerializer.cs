#nullable disable

using System.Xml;
using static Aid.Xlsx.SpreadsheetML.Xml.OpenXmlConst;
using static Aid.Xlsx.Model.Workbooks.XlsxWorkbook;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.RichText;
using Aid.Xlsx.Model.Styles.Formatting;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.SpreadsheetML.Strings;

internal class TextSerializer
{
    internal static void WriteRichTextElements(XmlWriter w, XlsxImmutableRichText richText, SaveContext context)
    {
        foreach (XlsxImmutableRichText.RichTextRun textRun in richText.Runs)
        {
            var text = richText.GetRunText(textRun);
            if (text.Length > 0)
            {
                WriteRun(w, text, textRun.Font);
            }
        }

        if (richText.PhoneticsProperties is not null)
        {
            XlsxImmutableRichText.PhoneticProperties phoneticsProps = richText.PhoneticsProperties.Value;
            foreach (XlsxImmutableRichText.PhoneticRun p in richText.PhoneticRuns)
            {
                w.WriteStartElement("rPh", Main2006SsNs);
                w.WriteAttribute("sb", p.StartIndex);
                w.WriteAttribute("eb", p.EndIndex);

                w.WriteStartElement("t", Main2006SsNs);
                if (p.Text.PreserveSpaces())
                    w.WritePreserveSpaceAttr();

                w.WriteString(p.Text);
                w.WriteEndElement(); // t
                w.WriteEndElement(); // rPh
            }

            XlsxFontFormatValue phoneticsFont = phoneticsProps.Font;
            var phoneticsFontId = context.GetFontId(phoneticsFont);

            w.WriteStartElement("phoneticPr", Main2006SsNs);
            w.WriteAttribute("fontId", phoneticsFontId);

            if (phoneticsProps.Alignment != XlsxPhoneticAlignment.Left)
                w.WriteAttributeString("alignment", phoneticsProps.Alignment.ToOpenXmlString());

            if (phoneticsProps.Type != XlsxPhoneticType.FullWidthKatakana)
                w.WriteAttributeString("type", phoneticsProps.Type.ToOpenXmlString());

            w.WriteEndElement(); // phoneticPr
        }
    }

    internal static void WriteRun(XmlWriter w, XlsxImmutableRichText richText, XlsxImmutableRichText.RichTextRun run)
    {
        var runText = richText.GetRunText(run);
        WriteRun(w, runText, run.Font);
    }

    private static void WriteRun(XmlWriter w, string text, XlsxFontFormatValue font)
    {
        XlsxFontFormatValue defaultFont = XlsxFontFormatValue.Default;
        w.WriteStartElement("r", Main2006SsNs);
        w.WriteStartElement("rPr", Main2006SsNs);

        if (font.Bold)
            w.WriteEmptyElement("b");

        if (font.Italic)
            w.WriteEmptyElement("i");

        if (font.Strikethrough)
            w.WriteEmptyElement("strike");

        // Three attributes are not stored/written:
        // * outline - doesn't do anything and likely only works in Word.
        // * condense - legacy compatibility setting for macs
        // * extend - legacy compatibility setting for pre-xlsx Excels
        // None have sensible descriptions.

        if (font.Shadow)
            w.WriteEmptyElement("shadow");

        if (font.Underline != defaultFont.Underline)
            WriteRunProperty(w, "u", font.Underline.ToOpenXmlString());

        WriteRunProperty(w, @"vertAlign", font.VerticalAlignment.ToOpenXmlString());
        WriteRunProperty(w, "sz", font.Size.Points);
        w.WriteColor("color", font.Color);
        WriteRunProperty(w, "rFont", font.Name.Text);
        WriteRunProperty(w, "family", (int)font.Family);

        if (font.Charset != defaultFont.Charset)
            WriteRunProperty(w, "charset", (int)font.Charset);

        if (font.Scheme != defaultFont.Scheme)
            WriteRunProperty(w, "scheme", font.Scheme.ToOpenXmlString());

        w.WriteEndElement(); // rPr

        w.WriteStartElement("t", Main2006SsNs);
        if (text.PreserveSpaces())
            w.WritePreserveSpaceAttr();

        w.WriteString(text);

        w.WriteEndElement(); // t
        w.WriteEndElement(); // r
    }

    private static void WriteRunProperty(XmlWriter w, string elName, string val)
    {
        w.WriteStartElement(elName, Main2006SsNs);
        w.WriteAttributeString("val", val);
        w.WriteEndElement();
    }

    private static void WriteRunProperty(XmlWriter w, string elName, int val)
    {
        w.WriteStartElement(elName, Main2006SsNs);
        w.WriteAttribute("val", val);
        w.WriteEndElement();
    }

    private static void WriteRunProperty(XmlWriter w, string elName, double val)
    {
        w.WriteStartElement(elName, Main2006SsNs);
        w.WriteAttribute("val", val);
        w.WriteEndElement();
    }
}
