using System.Xml;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.Infrastructure.Extensions;

internal static class XmlWriterExtensions
{
    extension(XmlWriter w)
    {
        public void WriteAttribute(String attrName, String value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttributeOptional(String attrName, String? value)
        {
            if (!string.IsNullOrEmpty(value))
                w.WriteAttribute(attrName, value);
        }

        public void WriteAttribute(String attrName, Int32 value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttribute(String attrName, UInt32 value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttributeOptional(String attrName, UInt32? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public void WriteAttributeOptional(String attrName, Int32? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public void WriteAttribute(String attrName, Double value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public void WriteAttribute(String attrName, Boolean value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value ? "1" : "0");
            w.WriteEndAttribute();
        }

        public void WriteAttributeDefault(String attrName, Boolean value, Boolean defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        public void WriteAttributeOptional(String attrName, Boolean? value)
        {
            if (value is not null)
                w.WriteAttribute(attrName, value.Value);
        }

        public void WriteAttributeDefault(String attrName, int value, int defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        public void WriteAttributeDefault(String attrName, uint value, uint defaultValue)
        {
            if (value != defaultValue)
                w.WriteAttribute(attrName, value);
        }

        /// <summary>
        /// Write date in a format <c>2015-01-01T00:00:00</c> (ignore kind).
        /// </summary>
        public void WriteAttribute(String attrName, DateTime value)
        {
            w.WriteStartAttribute(attrName);
            w.WriteValue(value.ToString("s"));
            w.WriteEndAttribute();
        }

        public void WriteAttribute(String attrName, String ns, Double value)
        {
            w.WriteStartAttribute(attrName, ns);
            w.WriteNumberValue(value);
            w.WriteEndAttribute();
        }

        public void WriteNumberValue(Double value)
        {
            // G17 will survive roundtrip to file and back
            w.WriteValue(value.ToInvariantString());
        }

        public void WritePreserveSpaceAttr()
        {
            w.WriteAttributeString("xml", "space", OpenXmlConst.Xml1998Ns, "preserve");
        }

        public void WriteEmptyElement(String elName)
        {
            w.WriteStartElement(elName, OpenXmlConst.Main2006SsNs);
            w.WriteEndElement();
        }

        public void WriteColor(String elName, XlsxColor xlColor, Boolean isDifferential = false)
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
