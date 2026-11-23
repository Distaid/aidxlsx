using System.Diagnostics;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Styles.Formatting;

namespace Aid.Xlsx.Model.RichText;

/// <summary>
/// A class for holding <see cref="XlsxRichText"/> in a <see cref="SharedStringTable"/>.
/// It's immutable (keys in reverse dictionary can't change) and more memory efficient
/// than mutable rich text.
/// </summary>
[DebuggerDisplay("{Text}")]
internal sealed class XlsxImmutableRichText : IEquatable<XlsxImmutableRichText>
{
    private readonly RichTextRun[] _runs;
    private readonly PhoneticRun[] _phoneticRuns;

    internal XlsxImmutableRichText(string text, RichTextRun[] runs, PhoneticRun[] phoneticRuns, PhoneticProperties? phoneticsProps)
    {
        Text = text;
        _runs = runs;
        _phoneticRuns = phoneticRuns;
        PhoneticsProperties = phoneticsProps;
    }

    /// <summary>
    /// A text of a whole rich text, without styling.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Individual rich text runs that make up the <see cref="Text"/>, in ascending order, non-overlapping.
    /// </summary>
    public IReadOnlyList<RichTextRun> Runs => _runs;

    /// <summary>
    /// All phonetics runs of rich text. Empty array, if no phonetic run. In ascending order, non-overlapping.
    /// </summary>
    public IReadOnlyList<PhoneticRun> PhoneticRuns => _phoneticRuns;

    /// <summary>
    /// Properties used to display phonetic runs.
    /// </summary>
    public PhoneticProperties? PhoneticsProperties { get; }

    public bool Equals(XlsxImmutableRichText? other)
    {
        if (other is null)
            return false;

        if (ReferenceEquals(this, other))
            return true;

        return Text == other.Text &&
               _runs.SequenceEqual(other._runs) &&
               _phoneticRuns.SequenceEqual(other._phoneticRuns) &&
               Nullable.Equals(PhoneticsProperties, other.PhoneticsProperties);
    }

    public override bool Equals(object? obj)
    {
        return Equals(obj as XlsxImmutableRichText);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = Text.GetHashCode();
            hashCode = (hashCode * 397) ^ PhoneticsProperties.GetHashCode();
            foreach (var phoneticRun in _phoneticRuns)
                hashCode = (hashCode * 397) ^ phoneticRun.GetHashCode();

            foreach (var run in _runs)
                hashCode = (hashCode * 397) ^ run.GetHashCode();

            return hashCode;
        }
    }

    internal string GetRunText(RichTextRun run) => Text.Substring(run.StartIndex, run.Length);

    /// <summary>
    /// Create an immutable rich text with same content as the original <paramref name="formattedText"/>.
    /// </summary>
    internal static XlsxImmutableRichText Create<T>(XlsxFormattedText<T> formattedText)
    {
        var text = formattedText.Text;
        var runs = new RichTextRun[formattedText.Count];
        var runIdx = 0;
        var charStartIdx = 0;
        foreach (var richString in formattedText)
        {
            runs[runIdx++] = new RichTextRun(richString.Font, charStartIdx, richString.Text.Length);
            charStartIdx += richString.Text.Length;
        }

        PhoneticRun[] phoneticRuns;
        PhoneticProperties? phoneticProps;
        if (formattedText.HasPhonetics)
        {
            var rtPhonetics = formattedText.Phonetics;
            phoneticRuns = new PhoneticRun[rtPhonetics.Count];
            var phoneticRunIdx = 0;
            var prevPhoneticEndIdx = 0;
            foreach (var phonetic in formattedText.Phonetics)
            {
                if (phonetic.Start >= text.Length)
                    throw new ArgumentException("Phonetic run start index must be within the text boundaries.");

                if (phonetic.End > text.Length)
                    throw new ArgumentException("Phonetic run end index must be at most length of a text.");

                if (phonetic.Start < prevPhoneticEndIdx)
                    throw new ArgumentException("Phonetic runs must be in ascending order and can't overlap.");

                phoneticRuns[phoneticRunIdx++] = new PhoneticRun(phonetic.Text, phonetic.Start, phonetic.End);
                prevPhoneticEndIdx = phonetic.End;
            }

            phoneticProps = new PhoneticProperties(formattedText.Phonetics);
        }
        else
        {
            phoneticRuns = [];
            phoneticProps = null;
        }

        return new XlsxImmutableRichText(text, runs, phoneticRuns, phoneticProps);
    }

    /// <summary>
    /// Make a new rich text with runs that use the <paramref name="font"/> as base font and
    /// use the <see cref="RichTextRun.Dxf"/> to create the final run font. This is a temporary
    /// method because rich text should store only the dxf. The API object should use the dxf
    /// and the cell format.
    /// </summary>
    internal XlsxImmutableRichText WithBaseFont(XlsxFontFormatValue font)
    {
        var adjustedRuns = new RichTextRun[_runs.Length];
        for (var i = 0; i < _runs.Length; ++i)
        {
            var run = _runs[i];
            var adjustedFont = font.AdjustWith(run.Dxf);
            adjustedRuns[i] = new RichTextRun(adjustedFont, run.StartIndex, run.Length);
        }

        return new XlsxImmutableRichText(Text, [.. adjustedRuns], _phoneticRuns, PhoneticsProperties);
    }

    internal readonly struct RichTextRun : IEquatable<RichTextRun>
    {
        internal readonly int StartIndex;
        internal readonly int Length;
        internal readonly XlsxFontFormatValue Font;

        // TODO: Refactor rich text run to use only dxf
        /// <summary>
        /// The Dxf member is used only during sheet load to resolve, otherwise it's completely
        /// ignored (no equals or hash code). The mutable rich text API should use the format of
        /// a cell and combine it with the Dxf. The text run shouldn't contain full font value,
        /// it takes the font of a cell and combines it with the dxf of the run.
        /// </summary>
        internal readonly XlsxDifferentialFontValue Dxf;

        internal RichTextRun(XlsxFontFormatValue font, int startIndex, int length)
            : this(font, XlsxDifferentialFontValue.Empty, startIndex, length)
        {
        }

        internal RichTextRun(XlsxDifferentialFontValue dxf, int startIndex, int length)
            : this(XlsxFontFormatValue.Default, dxf, startIndex, length)
        {
        }

        internal RichTextRun(XlsxFontFormatValue font, XlsxDifferentialFontValue dxf, int startIndex, int length)
        {
            Font = font;
            Dxf = dxf;
            StartIndex = startIndex;
            Length = length;
        }

        public bool Equals(RichTextRun other)
        {
            return StartIndex == other.StartIndex && Length == other.Length && Font.Equals(other.Font);
        }

        public override bool Equals(object? obj)
        {
            return obj is RichTextRun other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = StartIndex;
                hashCode = (hashCode * 397) ^ Length;
                hashCode = (hashCode * 397) ^ Font.GetHashCode();
                return hashCode;
            }
        }
    }

    /// <summary>
    /// Phonetic runs can't overlap and must be in order (i.e. start index must be ascending).
    /// </summary>
    internal readonly struct PhoneticRun
    {
        /// <summary>
        /// Text that is displayed above a segment indicating how should segment be read.
        /// </summary>
        internal readonly string Text;

        /// <summary>
        /// Starting index of displayed phonetic (first character is 0).
        /// </summary>
        internal readonly int StartIndex;

        /// <summary>
        /// End index, excluding (the last index is a length of the rich text).
        /// </summary>
        internal readonly int EndIndex;

        public PhoneticRun(string text, int startIndex, int endIndex)
        {
            if (text.Length == 0)
                throw new ArgumentException("Phonetic run text can't be empty.", nameof(text));

            if (startIndex < 0)
                throw new ArgumentException("Start index index must be greater than 0.", nameof(startIndex));

            if (startIndex >= endIndex)
                throw new ArgumentException("Start index must be less than end index.", nameof(endIndex));

            Text = text;
            StartIndex = startIndex;
            EndIndex = endIndex;
        }

        public bool Equals(PhoneticRun other)
        {
            return Text == other.Text && StartIndex == other.StartIndex && EndIndex == other.EndIndex;
        }

        public override bool Equals(object? obj)
        {
            return obj is PhoneticRun other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Text.GetHashCode();
                hashCode = (hashCode * 397) ^ StartIndex;
                hashCode = (hashCode * 397) ^ EndIndex;
                return hashCode;
            }
        }
    }

    /// <summary>
    /// Properties of phonetic runs. All phonetic runs of a rich text have same font and other properties.
    /// </summary>
    internal readonly struct PhoneticProperties
    {
        /// <summary>
        /// Font used for text of phonetic runs. All phonetic runs use same font. There can be no phonetic runs,
        /// but with specified font (e.g. the mutable API has only specified font, but no text yet).
        /// </summary>
        public readonly XlsxFontFormatValue Font;

        /// <summary>
        /// Type of phonetics. Default is <see cref="XlsxPhoneticType.FullWidthKatakana"/>
        /// </summary>
        public readonly XlsxPhoneticType Type;

        /// <summary>
        /// Alignment of phonetics. Default is <see cref="XlsxPhoneticAlignment.Left"/>
        /// </summary>
        public readonly XlsxPhoneticAlignment Alignment;

        internal PhoneticProperties(XlsxPhonetics rtPhonetics)
            : this(rtPhonetics.Font, rtPhonetics.Type, rtPhonetics.Alignment)
        {
        }

        internal PhoneticProperties(XlsxFontFormatValue font, XlsxPhoneticType type, XlsxPhoneticAlignment alignment)
        {
            Font = font;
            Type = type;
            Alignment = alignment;
        }

        public bool Equals(PhoneticProperties other)
        {
            return Font.Equals(other.Font) && Type == other.Type && Alignment == other.Alignment;
        }

        public override bool Equals(object? obj)
        {
            return obj is PhoneticProperties other && Equals(other);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Font.GetHashCode();
                hashCode = (hashCode * 397) ^ (int)Type;
                hashCode = (hashCode * 397) ^ (int)Alignment;
                return hashCode;
            }
        }
    }
}
