using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Opc;
using Aid.Xlsx.RichText;
using Aid.Xlsx.Styles;
using Aid.Xlsx.Utils;

namespace Aid.Xlsx.SpreadsheetML;

/// <summary>
/// A universal two-way mapper of string representation of an enum value in the OOXML to Aid.Xlsx enum.
/// </summary>
internal sealed class XmlToEnumMapper : IEnumMapper
{
    /// <summary>
    /// A collection of all maps. The key is enum type, the value is BiDictionary&lt;string,SomeEnum&gt;
    /// Value can't be typed due to generic limitations (no common ancestor).
    /// </summary>
    private readonly Dictionary<Type, object> _enumMaps;

    private static readonly Lazy<XmlToEnumMapper> LazyInstance = new(CreateSpreadsheetMapper);

    internal static XmlToEnumMapper Instance => LazyInstance.Value;

    private XmlToEnumMapper(Dictionary<Type, object> maps)
    {
        _enumMaps = maps;
    }

    public bool TryGetEnum<TEnum>(string text, out TEnum enumValue)
        where TEnum : struct, Enum
    {
        var enumMap = GetEnumMap<TEnum>();
        return enumMap.KeyToValue.TryGetValue(text, out enumValue);
    }

    public bool TryGetText<TEnum>(TEnum enumValue, [MaybeNullWhen(false)] out string text)
        where TEnum : struct, Enum
    {
        var enumMap = GetEnumMap<TEnum>();
        return enumMap.ValueToKey.TryGetValue(enumValue, out text);
    }

    private BiDictionary<string, TEnum> GetEnumMap<TEnum>()
        where TEnum : struct, Enum
    {
        return (BiDictionary<string, TEnum>)_enumMaps[typeof(TEnum)];
    }

    private static XmlToEnumMapper CreateSpreadsheetMapper()
    {
        var builder = new Builder();

        // ST_FontScheme
        builder.Add(new BiDictionary<string, XlsxFontScheme>
        {
            { "none", XlsxFontScheme.None },
            { "major", XlsxFontScheme.Major },
            { "minor", XlsxFontScheme.Minor },
        });

        // ST_UnderlineValues
        builder.Add(new BiDictionary<string, XlsxFontUnderlineValues>
        {
            { "double", XlsxFontUnderlineValues.Double },
            { "doubleAccounting", XlsxFontUnderlineValues.DoubleAccounting },
            { "none", XlsxFontUnderlineValues.None },
            { "single", XlsxFontUnderlineValues.Single },
            { "singleAccounting", XlsxFontUnderlineValues.SingleAccounting },
        });

        // ST_VerticalAlignRun
        builder.Add(new BiDictionary<string, XlsxFontVerticalTextAlignmentValues>
        {
            { "baseline", XlsxFontVerticalTextAlignmentValues.Baseline },
            { "subscript", XlsxFontVerticalTextAlignmentValues.Subscript },
            { "superscript", XlsxFontVerticalTextAlignmentValues.Superscript },
        });

        // ST_PatternType
        builder.Add(new BiDictionary<string, XlsxFillPatternValues>
        {
            { "none", XlsxFillPatternValues.None },
            { "solid", XlsxFillPatternValues.Solid },
            { "mediumGray", XlsxFillPatternValues.MediumGray },
            { "darkGray", XlsxFillPatternValues.DarkGray },
            { "lightGray", XlsxFillPatternValues.LightGray },
            { "darkHorizontal", XlsxFillPatternValues.DarkHorizontal },
            { "darkVertical", XlsxFillPatternValues.DarkVertical },
            { "darkDown", XlsxFillPatternValues.DarkDown },
            { "darkUp", XlsxFillPatternValues.DarkUp },
            { "darkGrid", XlsxFillPatternValues.DarkGrid },
            { "darkTrellis", XlsxFillPatternValues.DarkTrellis },
            { "lightHorizontal", XlsxFillPatternValues.LightHorizontal },
            { "lightVertical", XlsxFillPatternValues.LightVertical },
            { "lightDown", XlsxFillPatternValues.LightDown },
            { "lightUp", XlsxFillPatternValues.LightUp },
            { "lightGrid", XlsxFillPatternValues.LightGrid },
            { "lightTrellis", XlsxFillPatternValues.LightTrellis },
            { "gray125", XlsxFillPatternValues.Gray125 },
            { "gray0625", XlsxFillPatternValues.Gray0625 },
        });

        // ST_BorderStyle
        builder.Add(new BiDictionary<string, XlsxBorderStyleValues>
        {
            { "none", XlsxBorderStyleValues.None },
            { "thin", XlsxBorderStyleValues.Thin },
            { "medium", XlsxBorderStyleValues.Medium },
            { "dashed", XlsxBorderStyleValues.Dashed },
            { "dotted", XlsxBorderStyleValues.Dotted },
            { "thick", XlsxBorderStyleValues.Thick },
            { "double", XlsxBorderStyleValues.Double },
            { "hair", XlsxBorderStyleValues.Hair },
            { "mediumDashed", XlsxBorderStyleValues.MediumDashed },
            { "dashDot", XlsxBorderStyleValues.DashDot },
            { "mediumDashDot", XlsxBorderStyleValues.MediumDashDot },
            { "dashDotDot", XlsxBorderStyleValues.DashDotDot },
            { "mediumDashDotDot", XlsxBorderStyleValues.MediumDashDotDot },
            { "slantDashDot", XlsxBorderStyleValues.SlantDashDot },
        });

        // ST_HorizontalAlignment
        builder.Add(new BiDictionary<string, XlsxAlignmentHorizontalValues>
        {
            { "general", XlsxAlignmentHorizontalValues.General },
            { "left", XlsxAlignmentHorizontalValues.Left },
            { "center", XlsxAlignmentHorizontalValues.Center },
            { "right", XlsxAlignmentHorizontalValues.Right },
            { "fill", XlsxAlignmentHorizontalValues.Fill },
            { "justify", XlsxAlignmentHorizontalValues.Justify },
            { "centerContinuous", XlsxAlignmentHorizontalValues.CenterContinuous },
            { "distributed", XlsxAlignmentHorizontalValues.Distributed },
        });

        // ST_VerticalAlignment
        builder.Add(new BiDictionary<string, XlsxAlignmentVerticalValues>
        {
            { "top", XlsxAlignmentVerticalValues.Top },
            { "center", XlsxAlignmentVerticalValues.Center },
            { "bottom", XlsxAlignmentVerticalValues.Bottom },
            { "justify", XlsxAlignmentVerticalValues.Justify },
            { "distributed", XlsxAlignmentVerticalValues.Distributed },
        });

        // ST_GradientType
        builder.Add(new BiDictionary<string, XlsxGradientType>
        {
            { "linear", XlsxGradientType.Linear },
            { "path", XlsxGradientType.Path }
        });

        // ST_PhoneticType
        builder.Add(new BiDictionary<string, XlsxPhoneticType>
        {
            { "halfwidthKatakana", XlsxPhoneticType.HalfWidthKatakana },
            { "fullwidthKatakana", XlsxPhoneticType.FullWidthKatakana },
            { "Hiragana", XlsxPhoneticType.Hiragana },
            { "noConversion", XlsxPhoneticType.NoConversion },
        });

        // ST_PhoneticAlignment
        builder.Add(new BiDictionary<string, XlsxPhoneticAlignment>
        {
            { "noControl", XlsxPhoneticAlignment.NoControl },
            { "left", XlsxPhoneticAlignment.Left },
            { "center", XlsxPhoneticAlignment.Center },
            { "distributed", XlsxPhoneticAlignment.Distributed },
        });

        return builder.Build();
    }

    internal class Builder
    {
        private readonly Dictionary<Type, object> _maps = new();

        public Builder Add<T>(BiDictionary<string, T> map)
            where T : struct, Enum
        {
            _maps.Add(typeof(T), map);
            return this;
        }

        public XmlToEnumMapper Build()
        {
            return new XmlToEnumMapper(_maps);
        }
    }
}
