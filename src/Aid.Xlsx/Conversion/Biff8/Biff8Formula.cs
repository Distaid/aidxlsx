using System.Buffers.Binary;
using System.Globalization;
using System.Text;
using System.Xml;
using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Conversion.Biff8;

/// <summary>Converts supported BIFF8 expression tokens to an A1 formula.</summary>
internal static class Biff8Formula
{
    internal static string? Decode(ReadOnlySpan<byte> tokens, ReadOnlySpan<string?> sheetPrefixes)
    {
        // MS-XLS Rgce limits the actual token size, including its in-memory expansion, to 1800 bytes.
        if (tokens.IsEmpty || tokens.Length > 1800)
            return null;

        var expressions = new List<Expression>();
        while (!tokens.IsEmpty)
        {
            var token = tokens[0];
            tokens = tokens[1..];
            // Reference, value and array classes use identical payloads for the tokens handled here.
            if (token is >= 0x40 and < 0x80)
                token = (byte)((token & 0x1f) | 0x20);

            if (token is >= 0x03 and <= 0x0e)
            {
                if (expressions.Count < 2 || expressions[^1].Text.Length == 0 || expressions[^2].Text.Length == 0)
                    return null;
                var (separator, precedence) = token switch
                {
                    0x03 => ("+", 3),
                    0x04 => ("-", 3),
                    0x05 => ("*", 4),
                    0x06 => ("/", 4),
                    0x07 => ("^", 5),
                    0x08 => ("&", 2),
                    0x09 => ("<", 1),
                    0x0a => ("<=", 1),
                    0x0b => ("=", 1),
                    0x0c => (">=", 1),
                    0x0d => (">", 1),
                    _ => ("<>", 1)
                };
                var right = expressions[^1].WithPrecedence(precedence + 1);
                expressions.RemoveAt(expressions.Count - 1);
                var left = expressions[^1].WithPrecedence(precedence);
                expressions[^1] = new Expression(left + separator + right, precedence);
                continue;
            }

            string? text;
            var precedenceOfExpression = int.MaxValue;
            switch (token)
            {
                case 0x12: // Unary plus
                case 0x13: // Unary minus
                case 0x14: // Percent
                case 0x15: // Explicit parentheses
                    if (expressions.Count == 0 || expressions[^1].Text.Length == 0)
                        return null;
                    var operand = expressions[^1];
                    precedenceOfExpression = token switch { 0x14 => 6, 0x15 => int.MaxValue, _ => 7 };
                    text = token switch
                    {
                        0x12 => "+" + operand.WithPrecedence(precedenceOfExpression),
                        0x13 => "-" + operand.WithPrecedence(precedenceOfExpression),
                        0x14 => operand.WithPrecedence(precedenceOfExpression) + "%",
                        _ => "(" + operand.Text + ")"
                    };
                    expressions.RemoveAt(expressions.Count - 1);
                    break;
                case 0x16: // Missing optional argument
                    text = string.Empty;
                    break;
                case 0x17: // String constant
                    text = ReadString(ref tokens);
                    break;
                case 0x19: // Optimized function and layout attributes
                    if (tokens.Length < 3)
                        return null;
                    var attribute = tokens[0];
                    var offsetCount = ReadU16(tokens[1..]);
                    tokens = tokens[3..];
                    if (attribute == 0x10)
                    {
                        text = ApplyFunction(expressions, "SUM", 1);
                        break;
                    }
                    if (attribute == 0x04)
                    {
                        var jumpBytes = 2 * (offsetCount + 1);
                        if (jumpBytes > tokens.Length)
                            return null;
                        tokens = tokens[jumpBytes..];
                    }
                    else if (attribute is not (0x01 or 0x02 or 0x08 or 0x40 or 0x41))
                    {
                        return null;
                    }
                    continue;
                case 0x1c: // Error constant
                    if (tokens.IsEmpty)
                        return null;
                    text = tokens[0] switch
                    {
                        0 => "#NULL!",
                        7 => "#DIV/0!",
                        15 => "#VALUE!",
                        23 => "#REF!",
                        29 => "#NAME?",
                        36 => "#NUM!",
                        42 => "#N/A",
                        _ => null
                    };
                    tokens = tokens[1..];
                    break;
                case 0x1d: // Boolean constant
                    if (tokens.IsEmpty || tokens[0] > 1)
                        return null;
                    text = tokens[0] == 1 ? "TRUE" : "FALSE";
                    tokens = tokens[1..];
                    break;
                case 0x1e: // Integer constant
                    if (tokens.Length < 2)
                        return null;
                    text = ReadU16(tokens).ToString(CultureInfo.InvariantCulture);
                    tokens = tokens[2..];
                    break;
                case 0x1f: // Floating-point constant
                    if (tokens.Length < 8)
                        return null;
                    var number = BinaryPrimitives.ReadDoubleLittleEndian(tokens);
                    if (!double.IsFinite(number))
                        return null;
                    text = number.ToString("R", CultureInfo.InvariantCulture);
                    precedenceOfExpression = double.IsNegative(number) ? 7 : int.MaxValue;
                    tokens = tokens[8..];
                    break;
                case 0x21: // Fixed-arity function
                    if (tokens.Length < 2)
                        return null;
                    var (function, arguments) = Function(ReadU16(tokens));
                    text = arguments >= 0 ? ApplyFunction(expressions, function, arguments) : null;
                    tokens = tokens[2..];
                    break;
                case 0x22: // Variable-arity function
                    if (tokens.Length < 3)
                        return null;
                    text = ApplyFunction(expressions, Function(ReadU16(tokens[1..])).Name, tokens[0]);
                    tokens = tokens[3..];
                    break;
                case 0x24: // Cell reference
                    if (tokens.Length < 4)
                        return null;
                    text = Reference(ReadU16(tokens), ReadU16(tokens[2..]));
                    tokens = tokens[4..];
                    break;
                case 0x25: // Area reference
                    if (tokens.Length < 8)
                        return null;
                    var first = Reference(ReadU16(tokens), ReadU16(tokens[4..]));
                    var last = Reference(ReadU16(tokens[2..]), ReadU16(tokens[6..]));
                    text = first is not null && last is not null ? first + ":" + last : null;
                    tokens = tokens[8..];
                    break;
                case 0x3a: // Cell reference with a sheet qualifier
                case 0x3b: // Area reference with a sheet qualifier
                    var size = token == 0x3a ? 6 : 10;
                    if (tokens.Length < size)
                        return null;
                    var sheetIndex = ReadU16(tokens);
                    if (sheetIndex >= sheetPrefixes.Length || sheetPrefixes[sheetIndex] is not { } prefix)
                        return null;

                    if (token == 0x3a)
                    {
                        var address = Reference(ReadU16(tokens[2..]), ReadU16(tokens[4..]));
                        text = address is not null ? prefix + address : null;
                    }
                    else
                    {
                        var start = Reference(ReadU16(tokens[2..]), ReadU16(tokens[6..]));
                        var end = Reference(ReadU16(tokens[4..]), ReadU16(tokens[8..]));
                        text = start is not null && end is not null ? prefix + start + ":" + end : null;
                    }

                    tokens = tokens[size..];
                    break;
                default:
                    return null;
            }

            if (text is null || text.Length > 8192)
                return null;
            expressions.Add(new Expression(text, precedenceOfExpression));
        }

        return expressions.Count == 1 && expressions[0].Text.Length is > 0 and <= 8192 ? expressions[0].Text : null;
    }

    private static ushort ReadU16(ReadOnlySpan<byte> bytes) => BinaryPrimitives.ReadUInt16LittleEndian(bytes);

    private static string? ReadString(ref ReadOnlySpan<byte> tokens)
    {
        if (tokens.Length < 2 || tokens[1] > 1)
            return null;
        var length = tokens[0];
        var width = tokens[1] == 0 ? 1 : 2;
        var byteCount = length * width;
        if (tokens.Length < 2 + byteCount)
            return null;
        var encoded = tokens.Slice(2, byteCount);
        var text = width == 1 ? Encoding.Latin1.GetString(encoded) : Encoding.Unicode.GetString(encoded);
        tokens = tokens[(2 + byteCount)..];
        try
        {
            XmlConvert.VerifyXmlChars(text);
        }
        catch (XmlException)
        {
            return null;
        }
        return '"' + text.Replace("\"", "\"\"", StringComparison.Ordinal) + '"';
    }

    private static string? Reference(int row, int columnWithFlags)
    {
        // A standalone formula stores coordinates even if the reference is relative.
        var column = columnWithFlags & 0x3fff;
        if (column > 255)
            return null;
        var columnPrefix = (columnWithFlags & 0x4000) == 0 ? "$" : string.Empty;
        var rowPrefix = (columnWithFlags & 0x8000) == 0 ? "$" : string.Empty;
        return columnPrefix + XlsxHelper.GetColumnLetterFromNumber(column + 1) + rowPrefix + (row + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static string? ApplyFunction(List<Expression> expressions, string? name, int count)
    {
        if (name is null || count > expressions.Count || count > 30)
            return null;
        var start = expressions.Count - count;
        var formula = new StringBuilder(name).Append('(');
        for (var index = start; index < expressions.Count; index++)
        {
            if (index != start)
                formula.Append(',');
            formula.Append(expressions[index].Text);
        }
        expressions.RemoveRange(start, count);
        return formula.Append(')').ToString();
    }

    // Function IDs and fixed arity are defined in MS-XLS Ftab. Variable arity is read from PtgFuncVar.
    private static (string? Name, int FixedArity) Function(ushort id) => id switch
    {
        0 => ("COUNT", -1),
        1 => ("IF", -1),
        2 => ("ISNA", 1),
        3 => ("ISERROR", 1),
        4 => ("SUM", -1),
        5 => ("AVERAGE", -1),
        6 => ("MIN", -1),
        7 => ("MAX", -1),
        8 => ("ROW", -1),
        9 => ("COLUMN", -1),
        10 => ("NA", 0),
        15 => ("SIN", 1),
        16 => ("COS", 1),
        17 => ("TAN", 1),
        18 => ("ATAN", 1),
        19 => ("PI", 0),
        20 => ("SQRT", 1),
        21 => ("EXP", 1),
        22 => ("LN", 1),
        23 => ("LOG10", 1),
        24 => ("ABS", 1),
        25 => ("INT", 1),
        26 => ("SIGN", 1),
        27 => ("ROUND", 2),
        31 => ("MID", 3),
        32 => ("LEN", 1),
        33 => ("VALUE", 1),
        34 => ("TRUE", 0),
        35 => ("FALSE", 0),
        36 => ("AND", -1),
        37 => ("OR", -1),
        38 => ("NOT", 1),
        39 => ("MOD", 2),
        48 => ("TEXT", 2),
        63 => ("RAND", 0),
        65 => ("DATE", 3),
        66 => ("TIME", 3),
        67 => ("DAY", 1),
        68 => ("MONTH", 1),
        69 => ("YEAR", 1),
        71 => ("HOUR", 1),
        72 => ("MINUTE", 1),
        73 => ("SECOND", 1),
        74 => ("NOW", 0),
        97 => ("ATAN2", 2),
        98 => ("ASIN", 1),
        99 => ("ACOS", 1),
        100 => ("CHOOSE", -1),
        101 => ("HLOOKUP", -1),
        102 => ("VLOOKUP", -1),
        109 => ("LOG", -1),
        111 => ("CHAR", 1),
        112 => ("LOWER", 1),
        113 => ("UPPER", 1),
        114 => ("PROPER", 1),
        115 => ("LEFT", -1),
        116 => ("RIGHT", -1),
        117 => ("EXACT", 2),
        118 => ("TRIM", 1),
        119 => ("REPLACE", 4),
        120 => ("SUBSTITUTE", -1),
        121 => ("CODE", 1),
        124 => ("FIND", -1),
        126 => ("ISERR", 1),
        127 => ("ISTEXT", 1),
        128 => ("ISNUMBER", 1),
        129 => ("ISBLANK", 1),
        162 => ("CLEAN", 1),
        183 => ("PRODUCT", -1),
        184 => ("FACT", 1),
        190 => ("ISNONTEXT", 1),
        198 => ("ISLOGICAL", 1),
        212 => ("ROUNDUP", 2),
        213 => ("ROUNDDOWN", 2),
        221 => ("TODAY", 0),
        228 => ("SUMPRODUCT", -1),
        336 => ("CONCATENATE", -1),
        337 => ("POWER", 2),
        345 => ("SUMIF", -1),
        346 => ("COUNTIF", 2),
        347 => ("COUNTBLANK", 1),
        _ => (null, -1)
    };

    private readonly record struct Expression(string Text, int Precedence)
    {
        internal string WithPrecedence(int minimum) => Precedence < minimum ? "(" + Text + ")" : Text;
    }
}
