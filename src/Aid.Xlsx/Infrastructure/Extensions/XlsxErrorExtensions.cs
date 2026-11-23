using Aid.Xlsx.Calculation;

namespace Aid.Xlsx.Infrastructure.Extensions;

internal static class XlsxErrorExtensions
{
    public static string ToDisplayString(this XlsxError error) =>
        error switch
        {
            XlsxError.CellReference => "#REF!",
            XlsxError.IncompatibleValue => "#VALUE!",
            XlsxError.DivisionByZero => "#DIV/0!",
            XlsxError.NameNotRecognized => "#NAME?",
            XlsxError.NoValueAvailable => "#N/A",
            XlsxError.NullValue => "#NULL!",
            XlsxError.NumberInvalid => "#NUM!",
            _ => throw new ArgumentOutOfRangeException(nameof(error), error, "Unknown error value.")
        };
}

internal static class XlsxErrorParser
{
    private static readonly Dictionary<string, XlsxError> ErrorMap = new(StringComparer.Ordinal)
    {
        ["#REF!"] = XlsxError.CellReference,
        ["#VALUE!"] = XlsxError.IncompatibleValue,
        ["#DIV/0!"] = XlsxError.DivisionByZero,
        ["#NAME?"] = XlsxError.NameNotRecognized,
        ["#N/A"] = XlsxError.NoValueAvailable,
        ["#NULL!"] = XlsxError.NullValue,
        ["#NUM!"] = XlsxError.NumberInvalid
    };

    public static bool TryParseError(string input, out XlsxError error)
        => ErrorMap.TryGetValue(input.Trim(), out error);
}
