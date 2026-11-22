using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Conversion.Biff8;

internal sealed partial class Biff8WorkbookReader
{
    private readonly List<BiffCellFormat> _rawCellFormats = [];
    private readonly List<(XlsxFontFormatValue Font, ushort Color)> _fonts = [];
    private readonly Dictionary<ushort, string> _customFormats = [];
    private XlsxCellFormatValue?[] _cellFormats = [];
    private XlsxColor[] _palette = [];

    private void ReadFont(ReadOnlySpan<byte> record)
    {
        var height = U16(record);
        var flags = U16(record[2..]);
        var name = ReadCharacters(record[16..], record[14], (record[15] & 1) != 0);
        var font = XlsxFontFormatValue.Default with
        {
            Name = string.IsNullOrWhiteSpace(name) || name.Length > 31 ? XlsxFontFormatValue.Default.Name : (XlsxFontName)name,
            Size = height == 0 ? XlsxFontFormatValue.Default.Size : XlsxFontSize.FromPoints(height / 20d),
            Charset = (XlsxFontCharSet)record[12],
            Family = record[11] <= 5 ? (XlsxFontFamilyNumberingValues)record[11] : XlsxFontFamilyNumberingValues.NotApplicable,
            Bold = U16(record[6..]) >= 700,
            Italic = (flags & 0x02) != 0,
            Strikethrough = (flags & 0x08) != 0,
            Outline = (flags & 0x10) != 0,
            Shadow = (flags & 0x20) != 0,
            Condense = (flags & 0x40) != 0,
            Extend = (flags & 0x80) != 0,
            Underline = record[10] switch
            {
                1 => XlsxFontUnderlineValues.Single,
                2 => XlsxFontUnderlineValues.Double,
                0x21 => XlsxFontUnderlineValues.SingleAccounting,
                0x22 => XlsxFontUnderlineValues.DoubleAccounting,
                _ => XlsxFontUnderlineValues.None
            },
            VerticalAlignment = U16(record[8..]) switch
            {
                1 => XlsxFontVerticalTextAlignmentValues.Superscript,
                2 => XlsxFontVerticalTextAlignmentValues.Subscript,
                _ => XlsxFontVerticalTextAlignmentValues.Baseline
            }
        };
        _fonts.Add((font, U16(record[4..])));
    }

    private void ReadPalette(ReadOnlySpan<byte> record)
    {
        var count = U16(record);
        if (count > 56 || record.Length != 2 + 4 * count)
            throw new InvalidDataException("Invalid BIFF color palette.");

        _palette = new XlsxColor[count];
        for (var i = 0; i < count; i++)
        {
            var rgb = record[(2 + 4 * i)..];
            _palette[i] = XlsxColor.FromArgb(rgb[0], rgb[1], rgb[2]);
        }
    }

    private XlsxColor GetColor(int index)
    {
        var paletteIndex = index - 8;
        if ((uint)paletteIndex < (uint)_palette.Length)
            return _palette[paletteIndex];

        return XlsxColor.IndexedColors.GetValueOrDefault(index, XlsxColor.Automatic);
    }

    private void ReadCellFormat(ReadOnlySpan<byte> record)
    {
        _rawCellFormats.Add(new BiffCellFormat(U16(record), U16(record[2..]), U16(record[4..]),
            record[6], record[7], record[8], U32(record[10..]), U32(record[14..]), U16(record[18..])));
    }

    private XlsxCellFormatValue BuildCellFormat(XlsxWorkbookStyles styles, BiffCellFormat source)
    {
        // BIFF skips font index four. Each XF stores the current values, including values inherited from a named style.
        var fontIndex = source.Font > 4 ? source.Font - 1 : source.Font;
        var font = styles.DefaultFormat.Font;
        if (source.Font != 4 && fontIndex < _fonts.Count)
        {
            var (definition, color) = _fonts[fontIndex];
            font = definition with { Color = GetColor(color) };
        }

        var numberFormat = _customFormats.TryGetValue(source.NumberFormat, out var custom)
            ? XlsxNumberFormat.Parse(custom)
            : styles.NumberFormats.TryGetValue(source.NumberFormat, out var builtIn)
                ? builtIn
                : styles.DefaultFormat.NumberFormat;

        var rotation = source.Rotation switch
        {
            <= 90 or 255 => source.Rotation,
            <= 180 => 90 - source.Rotation,
            _ => 0
        };
        var alignment = XlsxAlignmentFormatValue.Default with
        {
            Horizontal = HorizontalAlignments[source.Alignment & 7],
            Vertical = VerticalAlignments[Math.Min((source.Alignment >> 4) & 7, 5)],
            WrapText = (source.Alignment & 0x08) != 0,
            JustifyLastLine = (source.Alignment & 0x80) != 0,
            TextRotation = new TextRotation(rotation),
            Indent = source.IndentAndReadingOrder & 0x0F,
            ShrinkToFit = (source.IndentAndReadingOrder & 0x10) != 0,
            ReadingOrder = (source.IndentAndReadingOrder >> 6) switch
            {
                1 => XlsxAlignmentReadingOrderValues.LeftToRight,
                2 => XlsxAlignmentReadingOrderValues.RightToLeft,
                _ => XlsxAlignmentReadingOrderValues.ContextDependent
            }
        };

        var lines = source.BorderLines;
        var extra = source.BorderAndFill;
        var border = new XlsxBorderFormatValue
        {
            Left = GetBorderLine(lines, 0, lines, 16),
            Right = GetBorderLine(lines, 4, lines, 23),
            Top = GetBorderLine(lines, 8, extra, 0),
            Bottom = GetBorderLine(lines, 12, extra, 7),
            Diagonal = GetBorderLine(extra, 21, extra, 14),
            DiagonalDown = (lines & 0x40000000) != 0,
            DiagonalUp = (lines & 0x80000000) != 0,
        };
        var patternIndex = extra >> 26;
        var fill = patternIndex > 0 && patternIndex < FillPatterns.Length
            ? new XlsxFillFormatValue(new XlsxPatternFill
            {
                PatternType = FillPatterns[patternIndex],
                PatternColor = GetColor(source.FillColors & 0x7F),
                BackgroundColor = GetColor((source.FillColors >> 7) & 0x7F)
            })
            : XlsxFillFormatValue.None;

        return styles.GetRegisteredCellFormat(styles.DefaultFormat with
        {
            Font = font,
            NumberFormat = numberFormat,
            Alignment = alignment,
            Border = border,
            Fill = fill,
            Protection = new XlsxProtectionFormatValue { Locked = (source.Flags & 1) != 0, Hidden = (source.Flags & 2) != 0 },
            IncludeQuotePrefix = (source.Flags & 8) != 0,
            CustomFormat = CellFormatComponents.All,
        });
    }

    private XlsxBorderLine GetBorderLine(uint styles, int styleShift, uint colors, int colorShift)
    {
        var code = (styles >> styleShift) & 0x0F;
        return code is > 0 and < 14
            ? new XlsxBorderLine(GetColor((int)((colors >> colorShift) & 0x7F)), BorderStyles[code])
            : XlsxBorderLine.None;
    }

    private static readonly XlsxAlignmentHorizontalValues[] HorizontalAlignments =
    [
        XlsxAlignmentHorizontalValues.General, XlsxAlignmentHorizontalValues.Left, XlsxAlignmentHorizontalValues.Center,
        XlsxAlignmentHorizontalValues.Right, XlsxAlignmentHorizontalValues.Fill, XlsxAlignmentHorizontalValues.Justify,
        XlsxAlignmentHorizontalValues.CenterContinuous, XlsxAlignmentHorizontalValues.Distributed
    ];

    private static readonly XlsxAlignmentVerticalValues[] VerticalAlignments =
    [
        XlsxAlignmentVerticalValues.Top, XlsxAlignmentVerticalValues.Center, XlsxAlignmentVerticalValues.Bottom,
        XlsxAlignmentVerticalValues.Justify, XlsxAlignmentVerticalValues.Distributed, XlsxAlignmentVerticalValues.Bottom
    ];

    private static readonly XlsxBorderStyleValues[] BorderStyles =
    [
        XlsxBorderStyleValues.None, XlsxBorderStyleValues.Thin, XlsxBorderStyleValues.Medium, XlsxBorderStyleValues.Dashed,
        XlsxBorderStyleValues.Dotted, XlsxBorderStyleValues.Thick, XlsxBorderStyleValues.Double, XlsxBorderStyleValues.Hair,
        XlsxBorderStyleValues.MediumDashed, XlsxBorderStyleValues.DashDot, XlsxBorderStyleValues.MediumDashDot,
        XlsxBorderStyleValues.DashDotDot, XlsxBorderStyleValues.MediumDashDotDot, XlsxBorderStyleValues.SlantDashDot
    ];

    private static readonly XlsxFillPatternValues[] FillPatterns =
    [
        XlsxFillPatternValues.None, XlsxFillPatternValues.Solid, XlsxFillPatternValues.MediumGray, XlsxFillPatternValues.DarkGray,
        XlsxFillPatternValues.LightGray, XlsxFillPatternValues.DarkHorizontal, XlsxFillPatternValues.DarkVertical,
        XlsxFillPatternValues.DarkDown, XlsxFillPatternValues.DarkUp, XlsxFillPatternValues.DarkGrid, XlsxFillPatternValues.DarkTrellis,
        XlsxFillPatternValues.LightHorizontal, XlsxFillPatternValues.LightVertical, XlsxFillPatternValues.LightDown,
        XlsxFillPatternValues.LightUp, XlsxFillPatternValues.LightGrid, XlsxFillPatternValues.LightTrellis,
        XlsxFillPatternValues.Gray125, XlsxFillPatternValues.Gray0625
    ];

    private readonly record struct BiffCellFormat(ushort Font, ushort NumberFormat, ushort Flags, byte Alignment, byte Rotation,
        byte IndentAndReadingOrder, uint BorderLines, uint BorderAndFill, ushort FillColors);
}
