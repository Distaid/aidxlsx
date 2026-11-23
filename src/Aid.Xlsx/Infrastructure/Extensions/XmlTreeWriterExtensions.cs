using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.Infrastructure.Extensions;

/// <summary>
/// Extension methods for <see cref="XmlTreeWriter"/>. Keep the original class reasonably clean
/// and place convenience method here.
/// </summary>
internal static class XmlTreeWriterExtensions
{
    extension(XmlTreeWriter xml)
    {
        /// <summary>
        /// Write a <c>CT_BooleanProperty</c> element.
        /// </summary>
        public void WriteBooleanProperty(string elName, bool value, string ns)
        {
            xml.WriteStartElement(elName, ns);
            xml.WriteAttributeDefault("val", value, true);
            xml.WriteEndElement();
        }

        /// <summary>
        /// Write <c>ST_UnsignedIntHex</c> attribute value (8 hex digits).
        /// </summary>
        public void WriteAttributeHex(string attrName, uint value)
        {
            xml.WriteAttribute(attrName, value.ToString("X8"));
        }

        public void WriteAttributeDefault(string attrName, bool value, bool defaultValue)
        {
            if (value != defaultValue)
                xml.WriteAttribute(attrName, value);
        }

        public void WriteAttributeDefault(string attrName, int value, int defaultValue)
        {
            if (value != defaultValue)
                xml.WriteAttribute(attrName, value);
        }

        public void WriteAttributeDefault(string attrName, uint value, uint defaultValue)
        {
            if (value != defaultValue)
                xml.WriteAttribute(attrName, value);
        }

        public void WriteAttributeDefault(string attrName, double value, double defaultValue)
        {
            if (!value.Equals(defaultValue))
                xml.WriteAttribute(attrName, value);
        }

        public void WriteAttributeDefault<TEnum>(string attrName, TEnum enumValue, TEnum defaultValue)
            where TEnum : struct, Enum
        {
            if (!enumValue.Equals(defaultValue))
                xml.WriteAttribute(attrName, enumValue);
        }

        public void WriteAttributeOptional(string attrName, int? value)
        {
            if (value is not null)
                xml.WriteAttribute(attrName, value.Value);
        }

        public void WriteColor(string elName, string ns, XlsxColor color, bool isDifferential = false)
        {
            xml.WriteStartElement(elName, ns);
            switch (color.ColorType)
            {
                case XlsxColorType.Automatic:
                    xml.WriteAttribute("auto", 1);
                    break;

                case XlsxColorType.Color:
                    xml.WriteAttribute("rgb", color.Color.ToHex());
                    break;

                case XlsxColorType.Indexed:
                    // 64 is 'transparent' and should be ignored for differential formats
                    if (!isDifferential || color.Indexed != 64)
                        xml.WriteAttribute("indexed", color.Indexed);
                    break;

                case XlsxColorType.Theme:
                    xml.WriteAttribute("theme", (int)color.ThemeColor);

                    if (color.ThemeTint != 0)
                        xml.WriteAttribute("tint", color.ThemeTint);
                    break;
            }

            xml.WriteEndElement();
        }
    }
}
