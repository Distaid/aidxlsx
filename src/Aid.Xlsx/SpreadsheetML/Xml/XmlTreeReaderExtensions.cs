using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Infrastructure.Utilities;
using Aid.Xlsx.Model.Grid.Coordinates;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Opc.Xml;

namespace Aid.Xlsx.SpreadsheetML.Xml;

/// <summary>
/// A helper methods for patterns and types commonly found in OOXML. Reading concrete types is not
/// something for <see cref="XmlTreeReader"/>.
/// </summary>
internal static class XmlTreeReaderExtensions
{
    extension(XmlTreeReader reader)
    {
        /// <summary>
        /// Try to open an optional <c>CT_Color</c> and read the color.
        /// </summary>
        /// <returns><c>true</c> when there was a color element, false if there wasn't.</returns>
        public bool TryReadColor(string colorElementName, string ns, [NotNullWhen(true)] out XlsxColor? color)
        {
            if (reader.TryOpen(colorElementName, ns))
            {
                color = reader.ParseColor(colorElementName, ns);
                return true;
            }

            color = null;
            return false;
        }

        /// <summary>
        /// Read <c>CT_Color</c>.
        /// </summary>
        public XlsxColor ParseColor(string colorElementName, string ns)
        {
            // OI-29500: Office prioritizes the attributes as auto < indexed < rgb < theme, and only
            // round trips the type with the highest priority if two or more are specified.
            var theme = reader.GetOptionalUInt("theme");
            if (theme is not null)
            {
                var tint = reader.GetOptionalDouble("tint") ?? 0;
                var themeColor = XlsxColor.FromTheme((XlsxThemeColor)theme.Value, tint);
                reader.Close(colorElementName, ns);
                return themeColor;
            }

            var rgb = reader.GetOptionalString("rgb");
            if (rgb is not null)
            {
                var rgbColor = XlsxColor.FromColor(ColorStringParser.ParseFromArgb(rgb.AsSpan()));
                reader.Close(colorElementName, ns);
                return rgbColor;
            }

            var indexed = reader.GetOptionalUintAsInt("indexed");
            if (indexed is not null)
            {
                var indexedColor = indexed <= XlsxColor.IndexedColors.Count ? XlsxColor.FromIndex(indexed.Value) : XlsxColor.Automatic;
                reader.Close(colorElementName, ns);
                return indexedColor;
            }

            // All attributes of CT_Color are optional. An empty color uses the automatic color.
            _ = reader.GetOptionalBool("auto");
            reader.Close(colorElementName, ns);
            return XlsxColor.Automatic;
        }

        /// <summary>
        /// Read <c>CT_BooleanProperty</c>.
        /// </summary>
        public bool TryReadBoolElement(string boolElementName, string ns, out bool value)
        {
            if (!reader.TryOpen(boolElementName, ns))
            {
                value = false;
                return false;
            }

            value = reader.GetOptionalBool("val") ?? true;

            reader.Close(boolElementName, ns);
            return true;
        }

        /// <summary>
        /// Try to read an element with an attribute <c>val</c> that contains a <c>XString</c>. Example:
        /// <![CDATA[
        /// <xsd:complexType name="CT_FontName">
        ///   <xsd:attribute name="val" type="s:ST_Xstring" use="required"/>
        /// </xsd:complexType>
        /// ]]>
        /// </summary>
        public bool TryReadXStringValElement(string elementName, string ns, [NotNullWhen(true)] out string? value)
        {
            if (reader.TryOpen(elementName, ns))
            {
                value = reader.GetXString("val");
                reader.Close(elementName, ns);
                return true;
            }

            value = null;
            return false;
        }

        /// <summary>
        /// Try to read an optional element of type <c>CT_IntProperty</c>
        /// </summary>
        public bool TryReadIntValElement(string elementName, string ns, out int value)
        {
            if (reader.TryOpen(elementName, ns))
            {
                value = reader.GetInt("val");
                reader.Close(elementName, ns);
                return true;
            }

            value = default;
            return false;
        }

        /// <summary>
        /// Try to read an optional element that contains a required enum in a <c>val</c> attribute.
        /// <![CDATA[
        /// <xsd:complexType name="CT_FontScheme">
        ///   <xsd:attribute name="val" type="ST_FontScheme" use="required"/>
        /// </xsd:complexType>
        /// ]]>
        /// </summary>
        /// <returns><c>true</c> when there was the element, <c>false</c> when there wasn't the element.</returns>
        public bool TryReadEnumValElement<TEnum>(string elementName, string ns, [NotNullWhen(true)] out TEnum? enumValue)
            where TEnum : struct, Enum
        {
            if (reader.TryOpen(elementName, ns))
            {
                enumValue = reader.GetEnum<TEnum>("val");
                reader.Close(elementName, ns);
                return true;
            }

            enumValue = null;
            return false;
        }

        /// <summary>
        /// Parse elements that have value in the content, not in the attributes (e.g., formulas).
        /// </summary>
        public Xpr<string> ParseXString(string elementName, string ns)
        {
            if (!reader.TryOpen(elementName, ns))
            {
                return Xpr.Fail<string>();
            }

            var content = XStringConvert.Decode(reader.GetContent());
            reader.Close(elementName, ns);
            return Xpr.From(content);
        }

        /// <summary>
        /// Return optional <c>ST_CellRef</c>.
        /// </summary>
        public Point? GetOptionalPoint(string attributeName)
        {
            var attributeValue = reader.GetOptionalString(attributeName);
            if (attributeValue is null)
                return null;

            if (!Point.TryParse(attributeValue, out var point))
            {
                if (reader.StrictAttributeParsing)
                    throw PartStructureException.InvalidAttributeFormat(attributeName, attributeValue, reader);

                return null;
            }

            return point;
        }

        /// <summary>
        /// Return optional <c>ST_Ref</c>.
        /// </summary>
        public Area? GetOptionalArea(string attributeName)
        {
            var attributeValue = reader.GetOptionalString(attributeName);
            if (attributeValue is null)
                return null;

            if (!Area.TryParse(attributeValue, out var area))
            {
                if (reader.StrictAttributeParsing)
                    throw PartStructureException.InvalidAttributeFormat(attributeName, attributeValue, reader);

                return null;
            }

            return area;
        }
    }
}
