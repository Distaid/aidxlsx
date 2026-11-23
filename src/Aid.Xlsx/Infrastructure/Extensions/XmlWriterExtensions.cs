using System.Xml;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.Infrastructure.Extensions;

internal static class XmlWriterExtensions
{
    extension(XmlWriter w)
    {
        public void WriteAttribute(string attrName, string value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttributeOptional(string attrName, string? value)
        {
            if (!string.IsNullOrEmpty(value))
                w.WriteAttribute(attrName, value);
        }

        public void WriteAttribute(string attrName, int value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttribute(string attrName, uint value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttributeOptional(string attrName, uint? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public void WriteAttributeOptional(string attrName, int? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public void WriteAttribute(string attrName, double value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttribute(string attrName, bool value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value ? "1" : "0");
            w.WriteEndAttribute();
        }

        public void WriteAttributeDefault(string attrName, bool value, bool defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        public void WriteAttributeOptional(string attrName, bool? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public void WriteAttributeDefault(string attrName, int value, int defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        public void WriteAttributeDefault(string attrName, uint value, uint defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        /// <summary>
        /// Write date in a format <c>2015-01-01T00:00:00</c> (ignore kind).
        /// </summary>
        public void WriteAttribute(string attrName, DateTime value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value.ToString("s"));
            w.WriteEndAttribute();
        }

        public void WriteAttribute(string attrName, string ns, double value)
        {
            w.WriteStartAttribute(attrName, ns);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public void WriteNumberValue(double value)
        {
            // G17 will survive roundtrip to file and back
            w.WriteValue(value.ToInvariantString());
        }

        public void WritePreserveSpaceAttr()
        {
            w.WriteAttributeString("xml", "space", OpenXmlConst.Xml1998Ns, "preserve");
        }

        public void WriteEmptyElement(string elName)
        {
            w.WriteStartElement(elName, OpenXmlConst.Main2006SsNs);
            w.WriteEndElement();
        }

        public void WriteColor(string elName, XlsxColor xlColor, bool isDifferential = false)
        {
            w.WriteStartElement(elName, OpenXmlConst.Main2006SsNs);
            switch (xlColor.ColorType)
            {
                case XlsxColorType.Color:
                    w.WriteAttributeString("rgb", xlColor.Color.ToHex());
                    break;

                case XlsxColorType.Indexed:
                    // 64 is 'transparent' and should be ignored for differential formats
                    if (!isDifferential || xlColor.Indexed != 64)
                        w.WriteAttribute("indexed", xlColor.Indexed);
                    break;

                case XlsxColorType.Theme:
                    w.WriteAttribute("theme", (int)xlColor.ThemeColor);

                    if (xlColor.ThemeTint != 0)
                        w.WriteAttribute("tint", xlColor.ThemeTint);
                    break;
            }

            w.WriteEndElement();
        }
    }
}
