#nullable disable

using System.Diagnostics;
using System.Globalization;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.Model.Grid.Cells;

/// <summary>
/// A value of a single cell. It contains a value a specific <see cref="Type"/>.
/// Structure provides following group of methods:
/// <list type="bullet">
///   <item><c>Is*</c> properties to check type (<see cref="IsNumber"/>, <see cref="IsBlank"/>...)</item>
///   <item><c>Get*</c> methods that return the stored value or throw <see cref="InvalidCastException"/> for incorrect type.</item>
///   <item>Explicit operators to convert <c>XlsxCellValue</c> to a concrete type. It is an equivalent of <c>Get*</c> methods.</item>
///   <item><c>TryConvert</c> methods to try to get value of a specific type, even if the value is of a different type.</item>
/// </list>
/// </summary>
[DebuggerDisplay("{Type} {_text != null ? (object)_text : (object)_value}")]
public readonly struct XlsxCellValue : IEquatable<XlsxCellValue>, IEquatable<Blank>, IEquatable<Boolean>, IEquatable<Double>, IEquatable<String>, IEquatable<XlsxError>, IEquatable<DateTime>, IEquatable<TimeSpan>, IEquatable<int>
{
    private readonly double _value;
    private readonly string _text;

    private XlsxCellValue(Blank _) : this()
    {
        Type = XlsxDataType.Blank;
    }

    private XlsxCellValue(bool logical) : this()
    {
        Type = XlsxDataType.Boolean;
        _value = logical ? 1d : 0d;
    }

    private XlsxCellValue(double number) : this()
    {
        if (Double.IsNaN(number) || Double.IsInfinity(number))
            throw new ArgumentException("Value can't be NaN or infinity.", nameof(number));

        Type = XlsxDataType.Number;
        _value = number;
    }

    private XlsxCellValue(string text) : this()
    {
        ArgumentNullException.ThrowIfNull(text);

        if (text.Length > 32767)
            throw new ArgumentOutOfRangeException(nameof(text), "Cells can hold a maximum of 32,767 characters.");

        Type = XlsxDataType.Text;
        _text = text;
    }

    private XlsxCellValue(XlsxError error) : this()
    {
        if (error < XlsxError.NullValue || error > XlsxError.NoValueAvailable)
            throw new ArgumentOutOfRangeException(nameof(error));

        Type = XlsxDataType.Error;
        _value = (double)error;
    }

    private XlsxCellValue(DateTime dateTime) : this()
    {
        Type = XlsxDataType.DateTime;
        _value = dateTime.ToSerialDateTime();
    }

    private XlsxCellValue(TimeSpan timeSpan) : this()
    {
        Type = XlsxDataType.TimeSpan;
        _value = timeSpan.ToSerialDateTime();
    }

    private XlsxCellValue(XlsxDataType type, double value) : this()
    {
        if (Double.IsNaN(value) || Double.IsInfinity(value))
            throw new ArgumentException("Value can't be NaN or infinity.", nameof(value));
        Type = type;
        _value = value;
    }

    /// <summary>
    /// Type of the value.
    /// </summary>
    public XlsxDataType Type { get; }

    /// <summary>
    /// Is the type of value <c>Blank</c>?
    /// </summary>
    public bool IsBlank => Type == XlsxDataType.Blank;

    /// <summary>
    /// Is the type of value <see cref="XlsxDataType.Boolean"/>?
    /// </summary>
    public bool IsBoolean => Type == XlsxDataType.Boolean;

    /// <summary>
    /// Is the type of value <see cref="XlsxDataType.Number"/>?
    /// </summary>
    public bool IsNumber => Type == XlsxDataType.Number;

    /// <summary>
    /// Is the type of value <see cref="XlsxDataType.Text"/>?
    /// </summary>
    public bool IsText => Type == XlsxDataType.Text;

    /// <summary>
    /// Is the type of value <see cref="XlsxDataType.Error"/>?
    /// </summary>
    public bool IsError => Type == XlsxDataType.Error;

    /// <summary>
    /// Is the type of value <see cref="XlsxDataType.DateTime"/>?
    /// </summary>
    public bool IsDateTime => Type == XlsxDataType.DateTime;

    /// <summary>
    /// Is the type of value <see cref="XlsxDataType.TimeSpan"/>?
    /// </summary>
    public bool IsTimeSpan => Type == XlsxDataType.TimeSpan;

    /// <summary>
    /// Is the value <see cref="XlsxDataType.Number"/> or <see cref="XlsxDataType.DateTime"/> or <see cref="XlsxDataType.TimeSpan"/>.
    /// </summary>
    public bool IsUnifiedNumber => IsNumber || IsDateTime || IsTimeSpan;

    public static implicit operator XlsxCellValue(Blank blank) => new(blank);
    public static implicit operator XlsxCellValue(bool logical) => new(logical);
    public static implicit operator XlsxCellValue(string text) => text is not null ? new(text) : new(Blank.Value);
    public static implicit operator XlsxCellValue(XlsxError error) => new(error);
    public static implicit operator XlsxCellValue(DateTime dateTime) => new(dateTime);
    public static implicit operator XlsxCellValue(TimeSpan timeSpan) => new(timeSpan);

    public static implicit operator XlsxCellValue(sbyte number) => new(number);
    public static implicit operator XlsxCellValue(byte number) => new(number);
    public static implicit operator XlsxCellValue(short number) => new(number);
    public static implicit operator XlsxCellValue(ushort number) => new(number);
    public static implicit operator XlsxCellValue(int number) => new(number);
    public static implicit operator XlsxCellValue(uint number) => new(number);
    public static implicit operator XlsxCellValue(long number) => new(number);
    public static implicit operator XlsxCellValue(ulong number) => new(number);
    public static implicit operator XlsxCellValue(float number) => new(number);
    public static implicit operator XlsxCellValue(double number) => new(number);
    public static implicit operator XlsxCellValue(decimal number) => new(decimal.ToDouble(number));

    public static implicit operator XlsxCellValue(sbyte? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(byte? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(short? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(ushort? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(int? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(uint? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(long? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(ulong? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(float? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(double? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(decimal? numberOrBlank) => numberOrBlank.HasValue ? numberOrBlank.Value : Blank.Value;

    public static implicit operator XlsxCellValue(DateTime? dateTimeOrBlank) => dateTimeOrBlank.HasValue ? dateTimeOrBlank.Value : Blank.Value;
    public static implicit operator XlsxCellValue(TimeSpan? timeSpanOrBlank) => timeSpanOrBlank.HasValue ? timeSpanOrBlank.Value : Blank.Value;

    /// <summary>
    /// Creates an <see cref="XlsxCellValue"/> from an <see cref="object"/>. If the type of the object has an implicit conversion operator then it is used.
    /// Otherwise, the <see cref="Convert.ToString(object, IFormatProvider)"/> method is used to convert the provided object to a string.
    /// <para/>
    /// The following types and their nullable counterparts are supported without requiring to be converted to a string:
    /// <list type="bullet">
    ///   <item><see cref="Blank"/></item>
    ///   <item><see cref="bool"/></item>
    ///   <item><see cref="string"/></item>
    ///   <item><see cref="XlsxError"/></item>
    ///   <item><see cref="DateTime"/></item>
    ///   <item><see cref="TimeSpan"/></item>
    ///   <item><see cref="sbyte"/></item>
    ///   <item><see cref="byte"/></item>
    ///   <item><see cref="short"/></item>
    ///   <item><see cref="ushort"/></item>
    ///   <item><see cref="int"/></item>
    ///   <item><see cref="uint"/></item>
    ///   <item><see cref="long"/></item>
    ///   <item><see cref="ulong"/></item>
    ///   <item><see cref="float"/></item>
    ///   <item><see cref="double"/></item>
    ///   <item><see cref="decimal"/></item>
    /// </list>
    /// </summary>
    /// <param name="obj">The object to convert.</param>
    /// <param name="provider">An object that supplies culture-specific formatting information.</param>
    /// <returns>An <see cref="XlsxCellValue"/> representation of the object.</returns>
    public static XlsxCellValue FromObject(object obj, IFormatProvider provider = null)
    {
        return obj switch
        {
            null => Blank.Value,
            Blank blank => blank,
            bool logical => logical,
            string text => text,
            XlsxError error => error,
            DateTime dateTime => dateTime,
            TimeSpan timeSpan => timeSpan,
            sbyte number => number,
            byte number => number,
            short number => number,
            ushort number => number,
            int number => number,
            uint number => number,
            long number => number,
            ulong number => number,
            float number => number,
            double number => number,
            decimal number => number,
            _ => Convert.ToString(obj, provider)
        };
    }

    /// <summary>
    /// A function used during data insertion to convert an <c>object</c> to <c>XlsxCellValue</c>.
    /// </summary>
    internal static XlsxCellValue FromInsertedObject(object value)
    {
        XlsxCellValue convertedValue = value switch
        {
            null => Blank.Value,
            Blank blankValue => blankValue,
            Boolean logical => logical,
            SByte number => number,
            Byte number => number,
            Int16 number => number,
            UInt16 number => number,
            Int32 number => number,
            UInt32 number => number,
            Int64 number => number,
            UInt64 number => number,
            Single number => number,
            Double number => number,
            Decimal number => number,
            String text => text,
            XlsxError error => error,
            DateTime date => date,
            DateTimeOffset dateOfs => dateOfs.DateTime,
            TimeSpan timeSpan => timeSpan,
            _ => value.ToString() // Other things, like chars ect are just turned to string
        };
        return convertedValue;
    }

    /// <summary>
    /// Try to convert a string into a string value, doing your best. If no other type can be
    /// extracted, consider it a text.
    /// </summary>
    /// <param name="text">Text to parse into a value.</param>
    /// <param name="culture">Culture used to parse numbers.</param>
    /// <returns>Parsed value.</returns>
    internal static XlsxCellValue FromText(string text, CultureInfo culture)
    {
        // AutoFilter custom filter operand can be stored as `1 1/2` and Excel correctly
        // interprets it as a `1.5`. Same for 2015-01-01, therefore use `TextToNumber` that
        // should deal with any weird formats.
        if (text is null)
            return Blank.Value;
        if (text == String.Empty)
            return Blank.Value;
        if (StringComparer.OrdinalIgnoreCase.Equals("TRUE", text))
            return true;
        if (StringComparer.OrdinalIgnoreCase.Equals("FALSE", text))
            return false;
        if (ScalarValue.TextToNumber(text, culture).TryPickT0(out var number, out _))
            return number;
        if (XlsxErrorParser.TryParseError(text, out var error))
            return error;

        return text;
    }

    /// <inheritdoc cref="GetBlank"/>
    public static explicit operator Blank(XlsxCellValue value) => value.GetBlank();

    /// <inheritdoc cref="GetBoolean"/>
    public static explicit operator Boolean(XlsxCellValue value) => value.GetBoolean();

    /// <inheritdoc cref="GetNumber"/>
    public static explicit operator Double(XlsxCellValue value) => value.GetNumber();

    /// <inheritdoc cref="GetText"/>
    public static explicit operator String(XlsxCellValue value) => value.GetText();

    /// <inheritdoc cref="GetError"/>
    public static explicit operator XlsxError(XlsxCellValue value) => value.GetError();

    /// <inheritdoc cref="GetDateTime"/>
    public static implicit operator DateTime(XlsxCellValue value) => value.GetDateTime();

    /// <inheritdoc cref="GetTimeSpan"/>
    public static implicit operator TimeSpan(XlsxCellValue value) => value.GetTimeSpan();

    /// <summary>
    /// If the value is of type <see cref="XlsxDataType.Blank"/>,
    /// return <see cref="Blank.Value"/>, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public Blank GetBlank() => IsBlank ? Blank.Value : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XlsxDataType.Boolean"/>,
    /// return logical, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public Boolean GetBoolean() => IsBoolean ? _value != 0d : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XlsxDataType.Number"/>,
    /// return number, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public Double GetNumber() => IsNumber ? _value : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XlsxDataType.Text"/>,
    /// return text, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public String GetText() => IsText ? _text : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XlsxDataType.Error"/>,
    /// return error, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public XlsxError GetError() => IsError ? (XlsxError)_value : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XlsxDataType.DateTime"/>,
    /// return date time, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public DateTime GetDateTime() => IsDateTime ? _value.ToSerialDateTime() : throw new InvalidCastException();

    /// <summary>
    /// If the value is of type <see cref="XlsxDataType.TimeSpan"/>,
    /// return time span, otherwise throw <see cref="InvalidCastException"/>.
    /// </summary>
    public TimeSpan GetTimeSpan() => IsTimeSpan ? _value.ToSerialTimeSpan() : throw new InvalidCastException();

    internal static XlsxCellValue FromSerialDateTime(double serialDateTime) => new(XlsxDataType.DateTime, serialDateTime);

    internal static XlsxCellValue FromSerialTimeSpan(double serialTimeSpan) => new(XlsxDataType.TimeSpan, serialTimeSpan);

    /// <summary>
    /// Get a number, either directly from type number or from serialized time span
    /// or serialized date time.
    /// </summary>
    /// <exception cref="InvalidCastException">If type is not <see cref="XlsxDataType.Number"/> or
    /// <see cref="XlsxDataType.DateTime"/> or <see cref="XlsxDataType.TimeSpan"/>.</exception>
    public double GetUnifiedNumber()
    {
        if (IsUnifiedNumber)
            return _value;

        throw new InvalidCastException("Value is not a number.");
    }

    internal Object ToObject()
    {
        return Type switch
        {
            XlsxDataType.Blank => null,
            XlsxDataType.Boolean => GetBoolean(),
            XlsxDataType.Number => GetNumber(),
            XlsxDataType.Text => GetText(),
            XlsxDataType.Error => GetError(),
            XlsxDataType.DateTime => GetDateTime(),
            XlsxDataType.TimeSpan => GetTimeSpan(),
            _ => throw new InvalidCastException()
        };
    }

    /// <summary>
    /// Return text representation of a value in current culture.
    /// </summary>
    public override string ToString() => ToString(CultureInfo.CurrentCulture);

    /// <summary>
    /// Return text representation of a value in the specified culture.
    /// </summary>
    public string ToString(CultureInfo culture) =>
        Type switch
        {
            XlsxDataType.Blank => string.Empty,
            XlsxDataType.Boolean => GetBoolean() ? "TRUE" : "FALSE",
            XlsxDataType.Number => _value.ToString(culture),
            XlsxDataType.Text => _text,
            XlsxDataType.Error => GetError().ToDisplayString(),
            XlsxDataType.DateTime => GetDateTime().ToString(culture),
            XlsxDataType.TimeSpan => GetTimeSpan().ToExcelString(culture),
            _ => throw new InvalidOperationException()
        };

    public bool Equals(XlsxCellValue other)
    {
        return Type == other.Type && _value.Equals(other._value) && _text == other._text;
    }

    public bool Equals(Blank other)
    {
        return IsBlank;
    }

    public bool Equals(bool other)
    {
        return IsBoolean && GetBoolean() == other;
    }

    public bool Equals(double other)
    {
        return IsNumber && _value.Equals(other);
    }

    /// <summary>
    /// Is the cell value text and is equal to the <paramref name="other"/>?
    /// Text comparison is case sensitive.
    /// </summary>
    public bool Equals(string other)
    {
        return IsText && _text == other;
    }

    public bool Equals(XlsxError other)
    {
        return IsError && GetError() == other;
    }

    public bool Equals(DateTime other)
    {
        return IsDateTime && GetDateTime() == other;
    }

    public bool Equals(TimeSpan other)
    {
        return IsTimeSpan && GetTimeSpan() == other;
    }

    public bool Equals(int other)
    {
        return Equals((double)other);
    }

    public override bool Equals(object obj)
    {
        return obj is XlsxCellValue other && Equals(other);
    }

    public override int GetHashCode()
    {
        unchecked
        {
            var hashCode = _value.GetHashCode();
            hashCode = (hashCode * 397) ^ (_text != null ? _text.GetHashCode() : 0);
            hashCode = (hashCode * 397) ^ (int)Type;
            return hashCode;
        }
    }

    /// <summary>
    /// Get a value, if it is a <see cref="XlsxDataType.Text"/>.
    /// </summary>
    /// <returns>True if value was retrieved, false otherwise.</returns>
    public bool TryGetText(out string value)
    {
        if (IsText)
        {
            value = _text;
            return true;
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="XlsxDataType.Blank"/> and return it.
    /// Method succeeds, when value is
    /// <list type="bullet">
    ///   <item>Type <see cref="XlsxDataType.Blank"/>.</item>
    ///   <item>Type <see cref="XlsxDataType.Text"/> and the text is empty.</item>
    /// </list>
    /// </summary>
    public bool TryConvert(out Blank value)
    {
        var isBlankLike = IsBlank || (IsText && GetText().Length == 0);
        if (isBlankLike)
        {
            value = Blank.Value;
            return true;
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="XlsxDataType.Boolean"/> and return it.
    /// Method succeeds, when value is
    /// <list type="bullet">
    ///   <item>Type <see cref="XlsxDataType.Boolean"/>.</item>
    ///   <item>Type <see cref="XlsxDataType.Number"/>, then the value of <c>0</c> means <c>false</c> and any other value is <c>true</c>.</item>
    ///   <item>Type <see cref="XlsxDataType.Text"/> and the value is <c>TRUE</c> or <c>FALSE</c> (case insensitive). Note that calc engine
    ///   generally doesn't coerce text to logical (e.g. arithmetic operations), though it happens in most function arguments (e.g.
    ///   <c>IF</c> or <c>AND</c>).</item>
    /// </list>
    /// </summary>
    public bool TryConvert(out Boolean value)
    {
        switch (Type)
        {
            case XlsxDataType.Boolean:
                value = GetBoolean();
                return true;
            case XlsxDataType.Number:
                value = GetNumber() != 0;
                return true;
            case XlsxDataType.Text when String.Equals(GetText(), "TRUE", StringComparison.OrdinalIgnoreCase):
                value = true;
                return true;
            case XlsxDataType.Text when String.Equals(GetText(), "FALSE", StringComparison.OrdinalIgnoreCase):
                value = false;
                return true;
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="XlsxDataType.Number"/> and return it.
    /// <list type="bullet">
    ///   <item>Double value - return the value.</item>
    ///   <item>Boolean value - return the <c>0</c> for <c>TRUE</c> and <c>1</c> for <c>FALSE</c>.</item>
    ///   <item>Text value - use the <c>VALUE</c> semantic to convert a text to a number.</item>
    ///   <item>DateTime value - return the serial date time number.</item>
    ///   <item>TimeSpan value - return the serial time span value.</item>
    /// </list>
    /// </summary>
    /// <remarks>Note that the coercion is current culture specific (e.g. decimal separators can differ).</remarks>
    /// <param name="value">The converted value. Result is never <c>infinity</c> or <c>NaN</c>.</param>
    /// <param name="culture">The culture used to convert the value for texts.</param>
    public bool TryConvert(out Double value, CultureInfo culture)
    {
        switch (Type)
        {
            case XlsxDataType.Number:
            case XlsxDataType.DateTime:
            case XlsxDataType.TimeSpan:
                value = GetUnifiedNumber();
                return true;
            case XlsxDataType.Boolean:
                value = GetBoolean() ? 1 : 0;
                return true;
            case XlsxDataType.Text:
                {
                    var coercionResult = ScalarValue.TextToNumber(GetText(), culture);
                    if (coercionResult.TryPickT0(out var number, out _))
                    {
                        value = number;
                        return true;
                    }

                    break;
                }
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="DateTime"/> and return it.
    /// Method succeeds, when value is
    /// <list type="bullet">
    ///   <item>Type <see cref="XlsxDataType.DateTime"/>.</item>
    ///   <item>Type <see cref="XlsxDataType.Number"/> and when the number is interpreted is a serial date time, it falls within the DateTime range.</item>
    ///   <item>Type <see cref="XlsxDataType.TimeSpan"/> and when the number is interpreted is a serial date time, it falls within the DateTime range.</item>
    /// </list>
    /// </summary>
    public bool TryConvert(out DateTime value)
    {
        if (IsUnifiedNumber)
        {
            var number = GetUnifiedNumber();
            if (number.IsValidOADateNumber())
            {
                value = number.ToSerialDateTime();
                return true;
            }
        }

        value = default;
        return false;
    }

    /// <summary>
    /// Try to convert the value to a <see cref="TimeSpan"/> and return it.
    /// Method succeeds, when value is
    /// <list type="bullet">
    ///   <item>Type <see cref="XlsxDataType.TimeSpan"/>.</item>
    ///   <item>Type <see cref="XlsxDataType.Number"/>, the number is interpreted is a time span date time.</item>
    ///   <item>Type <see cref="XlsxDataType.Text"/> and it can be parsed as a time span (accepts even hours over 24 hours).</item>
    /// </list>
    /// </summary>
    /// <param name="value">The converted value.</param>
    /// <param name="culture">The culture used to get time and decimal separators.</param>
    public bool TryConvert(out TimeSpan value, CultureInfo culture)
    {
        if (IsTimeSpan)
        {
            value = GetTimeSpan();
            return true;
        }

        if (IsNumber)
        {
            value = GetUnifiedNumber().ToSerialTimeSpan();
            return true;
        }

        if (IsText && TimeSpanParser.TryParseTime(GetText(), culture, out var ts))
        {
            value = ts;
            return true;
        }

        value = default;
        return false;
    }
}
