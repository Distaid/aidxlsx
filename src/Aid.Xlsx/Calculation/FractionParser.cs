#nullable disable

using System.Globalization;
using System.Text.RegularExpressions;

namespace Aid.Xlsx.Calculation;

/// <summary>
/// Parse a fraction for text-to-number type coercion.
/// </summary>
internal static partial class FractionParser
{
    [GeneratedRegex(@"^ *([+-]?) *([0-9]+) ([0-9]{1,5})/([0-9]{1,5}) *$", RegexOptions.CultureInvariant)]
    private static partial Regex FractionRegex { get; }

    public static bool TryParse(string s, out double result)
    {
        result = default;
        var match = FractionRegex.Match(s);
        if (!match.Success)
            return false;

        var denominator = ParseInt(match.Groups[4]);
        if (denominator == 0 || denominator > short.MaxValue)
            return false;

        var numerator = ParseInt(match.Groups[3]);
        if (numerator > short.MaxValue)
            return false;

        var sign = match.Groups[1];
        var wholeNumber = ParseInt(match.Groups[2]);

        var fraction = wholeNumber + numerator / (double)denominator;
        var hasNegativeSign = sign.Success && sign.Value.Length > 0 && sign.Value[0] == '-';
        result = hasNegativeSign ? -fraction : fraction;
        return true;

        static int ParseInt(Capture capture) =>
            int.Parse(capture.Value, NumberStyles.None, CultureInfo.InvariantCulture);
    }
}
