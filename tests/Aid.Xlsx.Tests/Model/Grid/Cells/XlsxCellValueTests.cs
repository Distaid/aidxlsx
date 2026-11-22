using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using NUnit.Framework;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.Tests.Model.Grid.Cells;

[TestFixture]
public class XlsxCellValueTests
{
    [Test]
    public void Creation_Blank()
    {
        XlsxCellValue blank = Blank.Value;
        Assert.AreEqual(XlsxDataType.Blank, blank.Type);
        Assert.True(blank.IsBlank);
    }

    [Test]
    public void Creation_Boolean()
    {
        XlsxCellValue logical = true;
        Assert.AreEqual(XlsxDataType.Boolean, logical.Type);
        Assert.True(logical.GetBoolean());
        Assert.True(logical.IsBoolean);
    }

    [Test]
    public void Creation_Number()
    {
        XlsxCellValue number = 14.0;
        Assert.AreEqual(XlsxDataType.Number, number.Type);
        Assert.True(number.IsNumber);
        Assert.AreEqual(14.0, number.GetNumber());
    }

    [TestCase(Double.NaN)]
    [TestCase(Double.PositiveInfinity)]
    [TestCase(Double.NegativeInfinity)]
    public void Creation_Number_CantBeNonNumber(Double nonNumber)
    {
        Assert.Throws<ArgumentException>(() => _ = (XlsxCellValue)nonNumber);
    }

    // Decimal is not allowed as a member of an attribute, so TestCase can't be used.
    private static readonly object[] DecimalTestCases =
    {
        new object[] { 5.875m, 5.875d },
        new object[] { Decimal.MaxValue, 7.922816251426434E+28 },
        new object[] { 1.0E-28m, 1.0000000000000001E-28d }
    };

    [TestCaseSource(nameof(DecimalTestCases))]
    public void Creation_Decimal(Decimal decimalNumber, Double expectedNumber)
    {
        XlsxCellValue cellValue = decimalNumber;
        Assert.True(cellValue.IsNumber);
        Assert.AreEqual(expectedNumber, cellValue.GetNumber());
    }

    [Test]
    public void Creation_Text()
    {
        XlsxCellValue text = "Hello World";
        Assert.AreEqual(XlsxDataType.Text, text.Type);
        Assert.AreEqual("Hello World", text.GetText());
    }

    [Test]
    public void NullString_IsConvertedToBlank()
    {
        XlsxCellValue value = (string)null;
        Assert.IsTrue(value.IsBlank);
        Assert.IsFalse(value.IsText);
    }

    [Test]
    public void Creation_Text_HasLimitedLength()
    {
        var longText = new string('A', 32768);
        Assert.Throws<ArgumentOutOfRangeException>(() => _ = (XlsxCellValue)longText);
    }

    [Test]
    public void Creation_Error()
    {
        XlsxCellValue error = XlsxError.NumberInvalid;
        Assert.AreEqual(XlsxDataType.Error, error.Type);
        Assert.True(error.IsError);
        Assert.AreEqual(XlsxError.NumberInvalid, error.GetError());
    }

    [Test]
    public void Creation_DateTime()
    {
        XlsxCellValue dateTime = new DateTime(2021, 1, 1);
        Assert.AreEqual(XlsxDataType.DateTime, dateTime.Type);
        Assert.True(dateTime.IsDateTime);
        Assert.AreEqual(new DateTime(2021, 1, 1), dateTime.GetDateTime());
    }

    [Test]
    public void Creation_TimeSpan()
    {
        XlsxCellValue dateTime = new TimeSpan(10, 1, 2, 3, 456);
        Assert.AreEqual(XlsxDataType.TimeSpan, dateTime.Type);
        Assert.True(dateTime.IsTimeSpan);
        Assert.AreEqual(new TimeSpan(10, 1, 2, 3, 456), dateTime.GetTimeSpan());
    }

    [Test]
    public void Creation_FromObject()
    {
        Assert.AreEqual(XlsxDataType.Blank, XlsxCellValue.FromObject(null).Type);
        Assert.AreEqual(XlsxDataType.Blank, XlsxCellValue.FromObject(Blank.Value).Type);
        Assert.AreEqual(XlsxDataType.Boolean, XlsxCellValue.FromObject(true).Type);
        Assert.AreEqual(XlsxDataType.Text, XlsxCellValue.FromObject("Hello World").Type);
        Assert.AreEqual(XlsxDataType.Error, XlsxCellValue.FromObject(XlsxError.NumberInvalid).Type);
        Assert.AreEqual(XlsxDataType.DateTime, XlsxCellValue.FromObject(new DateTime(2021, 1, 1)).Type);
        Assert.AreEqual(XlsxDataType.TimeSpan, XlsxCellValue.FromObject(new TimeSpan(10, 1, 2, 3, 456)).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((sbyte)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((byte)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((short)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((ushort)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((int)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((uint)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((long)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((ulong)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((float)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((double)42).Type);
        Assert.AreEqual(XlsxDataType.Number, XlsxCellValue.FromObject((decimal)42).Type);
        Assert.AreEqual(XlsxDataType.Text, XlsxCellValue.FromObject(DayOfWeek.Sunday).Type);
    }

    [Test]
    public void NumberTypes_HaveUnambiguousConversion()
    {
        {
            sbyte sbyteNumber = 5;
            XlsxCellValue sbyteCellValue = sbyteNumber;
            Assert.IsTrue(sbyteCellValue.IsNumber);
            Assert.AreEqual(5d, sbyteCellValue.GetNumber());
        }
        {
            byte byteNumber = 6;
            XlsxCellValue byteCellValue = byteNumber;
            Assert.IsTrue(byteCellValue.IsNumber);
            Assert.AreEqual(6d, byteCellValue.GetNumber());
        }
        {
            short shortNumber = 7;
            XlsxCellValue shortCellValue = shortNumber;
            Assert.IsTrue(shortCellValue.IsNumber);
            Assert.AreEqual(7d, shortCellValue.GetNumber());
        }
        {
            ushort ushortNumber = 8;
            XlsxCellValue ushortCellValue = ushortNumber;
            Assert.IsTrue(ushortCellValue.IsNumber);
            Assert.AreEqual(8d, ushortCellValue.GetNumber());
        }
        {
            int intNumber = 9;
            XlsxCellValue intCellValue = intNumber;
            Assert.IsTrue(intCellValue.IsNumber);
            Assert.AreEqual(9d, intCellValue.GetNumber());
        }
        {
            uint uintNumber = 10;
            XlsxCellValue uintCellValue = uintNumber;
            Assert.IsTrue(uintCellValue.IsNumber);
            Assert.AreEqual(10d, uintCellValue.GetNumber());
        }
        {
            long longNumber = 11;
            XlsxCellValue longCellValue = longNumber;
            Assert.IsTrue(longCellValue.IsNumber);
            Assert.AreEqual(11d, longCellValue.GetNumber());
        }
        {
            ulong ulongNumber = 12;
            XlsxCellValue ulongCellValue = ulongNumber;
            Assert.IsTrue(ulongCellValue.IsNumber);
            Assert.AreEqual(12d, ulongCellValue.GetNumber());
        }
        {
            float floatNumber = 13.5f;
            XlsxCellValue floatCellValue = floatNumber;
            Assert.IsTrue(floatCellValue.IsNumber);
            Assert.AreEqual(13.5d, floatCellValue.GetNumber());
        }
        {
            double doubleNumber = 14.5;
            XlsxCellValue doubleCellValue = doubleNumber;
            Assert.IsTrue(doubleCellValue.IsNumber);
            Assert.AreEqual(14.5d, doubleCellValue.GetNumber());
        }
        {
            decimal decimalNumber = 15.75m;
            XlsxCellValue decimalCellValue = decimalNumber;
            Assert.IsTrue(decimalCellValue.IsNumber);
            Assert.AreEqual(15.75d, decimalCellValue.GetNumber());
        }
    }

    [Test]
    [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
    public void NullableNumber_WithNullValue_AreConvertedToBlank()
    {
        {
            sbyte? sbyteNull = null;
            XlsxCellValue sbyteCellValue = sbyteNull;
            Assert.IsFalse(sbyteCellValue.IsNumber);
            Assert.IsTrue(sbyteCellValue.IsBlank);
        }
        {
            byte? byteNull = null;
            XlsxCellValue byteCellValue = byteNull;
            Assert.IsFalse(byteCellValue.IsNumber);
            Assert.IsTrue(byteCellValue.IsBlank);
        }
        {
            short? shortNull = null;
            XlsxCellValue shortCellValue = shortNull;
            Assert.IsFalse(shortCellValue.IsNumber);
            Assert.IsTrue(shortCellValue.IsBlank);
        }
        {
            ushort? ushortNull = null;
            XlsxCellValue ushortCellValue = ushortNull;
            Assert.IsFalse(ushortCellValue.IsNumber);
            Assert.IsTrue(ushortCellValue.IsBlank);
        }
        {
            int? intNull = null;
            XlsxCellValue intCellValue = intNull;
            Assert.IsFalse(intCellValue.IsNumber);
            Assert.IsTrue(intCellValue.IsBlank);
        }
        {
            uint? uintNull = null;
            XlsxCellValue uintCellValue = uintNull;
            Assert.IsFalse(uintCellValue.IsNumber);
            Assert.IsTrue(uintCellValue.IsBlank);
        }
        {
            long? longNull = null;
            XlsxCellValue longCellValue = longNull;
            Assert.IsFalse(longCellValue.IsNumber);
            Assert.IsTrue(longCellValue.IsBlank);
        }
        {
            ulong? ulongNull = null;
            XlsxCellValue ulongCellValue = ulongNull;
            Assert.IsFalse(ulongCellValue.IsNumber);
            Assert.IsTrue(ulongCellValue.IsBlank);
        }
        {
            float? floatValue = null;
            XlsxCellValue floatCellValue = floatValue;
            Assert.IsFalse(floatCellValue.IsNumber);
            Assert.IsTrue(floatCellValue.IsBlank);
        }
        {
            double? doubleValue = null;
            XlsxCellValue doubleCellValue = doubleValue;
            Assert.IsFalse(doubleCellValue.IsNumber);
            Assert.IsTrue(doubleCellValue.IsBlank);
        }
        {
            decimal? decimalValue = null;
            XlsxCellValue decimalCellValue = decimalValue;
            Assert.IsFalse(decimalCellValue.IsNumber);
            Assert.IsTrue(decimalCellValue.IsBlank);
        }
    }

    [Test]
    public void NullableNumber_WithNumberValue_AreConvertedToNumber()
    {
        {
            sbyte? sbyteNumber = 5;
            XlsxCellValue sbyteCellValue = sbyteNumber;
            Assert.IsTrue(sbyteCellValue.IsNumber);
            Assert.AreEqual(5d, sbyteCellValue.GetNumber());
        }
        {
            byte? byteNumber = 6;
            XlsxCellValue byteCellValue = byteNumber;
            Assert.IsTrue(byteCellValue.IsNumber);
            Assert.AreEqual(6d, byteCellValue.GetNumber());
        }
        {
            short? shortNumber = 7;
            XlsxCellValue shortCellValue = shortNumber;
            Assert.IsTrue(shortCellValue.IsNumber);
            Assert.AreEqual(7d, shortCellValue.GetNumber());
        }
        {
            ushort? ushortNumber = 8;
            XlsxCellValue ushortCellValue = ushortNumber;
            Assert.IsTrue(ushortCellValue.IsNumber);
            Assert.AreEqual(8d, ushortCellValue.GetNumber());
        }
        {
            int? intNumber = 9;
            XlsxCellValue intCellValue = intNumber;
            Assert.IsTrue(intCellValue.IsNumber);
            Assert.AreEqual(9d, intCellValue.GetNumber());
        }
        {
            uint? uintNumber = 9;
            XlsxCellValue uintCellValue = uintNumber;
            Assert.IsTrue(uintCellValue.IsNumber);
            Assert.AreEqual(9d, uintCellValue.GetNumber());
        }
        {
            long? longNumber = 10;
            XlsxCellValue longCellValue = longNumber;
            Assert.IsTrue(longCellValue.IsNumber);
            Assert.AreEqual(10d, longCellValue.GetNumber());
        }
        {
            ulong? ulongNumber = 11;
            XlsxCellValue ulongCellValue = ulongNumber;
            Assert.IsTrue(ulongCellValue.IsNumber);
            Assert.AreEqual(11d, ulongCellValue.GetNumber());
        }
        {
            float? floatNumber = 12.875f;
            XlsxCellValue floatCellValue = floatNumber;
            Assert.IsTrue(floatCellValue.IsNumber);
            Assert.AreEqual(12.875d, floatCellValue.GetNumber());
        }
        {
            double? doubleNumber = 13.875d;
            XlsxCellValue doubleCellValue = doubleNumber;
            Assert.IsTrue(doubleCellValue.IsNumber);
            Assert.AreEqual(13.875d, doubleCellValue.GetNumber());
        }
        {
            decimal? decimalNumber = 14.875m;
            XlsxCellValue decimalCellValue = decimalNumber;
            Assert.IsTrue(decimalCellValue.IsNumber);
            Assert.AreEqual(14.875d, decimalCellValue.GetNumber());
        }
    }

    [Test]
    [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
    public void NullableDateTime_WithNullValue_IsConvertedToBlank()
    {
        DateTime? dateTimeNull = null;
        XlsxCellValue dateTimeCellValue = dateTimeNull;
        Assert.IsFalse(dateTimeCellValue.IsDateTime);
        Assert.IsTrue(dateTimeCellValue.IsBlank);
    }

    [Test]
    public void NullableDateTime_WithDateValue_IsConvertedToDateTime()
    {
        DateTime? dateTime = new DateTime(2020, 5, 14, 8, 14, 30);
        XlsxCellValue dateTimeCellValue = dateTime;
        Assert.IsTrue(dateTimeCellValue.IsDateTime);
        Assert.AreEqual(dateTime.Value, dateTimeCellValue.GetDateTime());
    }

    [Test]
    [SuppressMessage("ReSharper", "ExpressionIsAlwaysNull")]
    public void NullableTimeSpan_WithNullValue_IsConvertedToBlank()
    {
        TimeSpan? timeSpanNull = null;
        XlsxCellValue timeSpanCellValue = timeSpanNull;
        Assert.IsFalse(timeSpanCellValue.IsTimeSpan);
        Assert.IsTrue(timeSpanCellValue.IsBlank);
    }

    [Test]
    public void NullableTimeSpan_WithTimeSpanValue_IsConvertedToTimeSpan()
    {
        TimeSpan? timeSpan = new TimeSpan(48, 12, 45, 30);
        XlsxCellValue timeSpanCellValue = timeSpan;
        Assert.IsTrue(timeSpanCellValue.IsTimeSpan);
        Assert.AreEqual(timeSpan.Value, timeSpanCellValue.GetTimeSpan());
    }

    [Test]
    public void UnifiedNumber_IsFormOf_Number_DateTime_And_TimeSpan()
    {
        XlsxCellValue value = Blank.Value;
        Assert.False(value.IsUnifiedNumber);

        value = true;
        Assert.False(value.IsUnifiedNumber);

        value = 14;
        Assert.True(value.IsUnifiedNumber);
        Assert.AreEqual(14.0, value.GetUnifiedNumber());

        value = new DateTime(1900, 1, 1);
        Assert.True(value.IsUnifiedNumber);
        Assert.AreEqual(1.0, value.GetUnifiedNumber());

        value = new TimeSpan(2, 12, 0, 0);
        Assert.True(value.IsUnifiedNumber);
        Assert.AreEqual(2.5, value.GetUnifiedNumber());

        value = "Text";
        Assert.False(value.IsUnifiedNumber);

        value = XlsxError.CellReference;
        Assert.False(value.IsUnifiedNumber);
    }

    [TestCase("1900-01-01", 1)]
    [TestCase("1900-01-02", 2)]
    [TestCase("1900-02-01", 32)]
    [TestCase("1900-02-28", 59)] // Excel assumes 1900 was a leap year and 29.1.1900 existed
    [TestCase("1900-03-01", 61)]
    [TestCase("2017-01-01", 42736)]
    public void SerialDateTime(string dateString, double expectedSerial)
    {
        XlsxCellValue date = DateTime.Parse(dateString);
        Assert.AreEqual(expectedSerial, date.GetUnifiedNumber());
    }

    [Test]
    [SetCulture("cs-CZ")]
    public void ToString_RespectsCulture()
    {
        XlsxCellValue v = Blank.Value;
        Assert.AreEqual(String.Empty, v.ToString());

        v = true;
        Assert.AreEqual("TRUE", v.ToString());

        v = 25.4;
        Assert.AreEqual("25,4", v.ToString());

        v = "Hello";
        Assert.AreEqual("Hello", v.ToString());

        v = XlsxError.IncompatibleValue;
        Assert.AreEqual("#VALUE!", v.ToString());

        v = new DateTime(1900, 1, 2);
        Assert.AreEqual("02.01.1900 0:00:00", v.ToString());

        v = new DateTime(1900, 3, 1, 4, 10, 5);
        Assert.AreEqual("01.03.1900 4:10:05", v.ToString());

        v = new TimeSpan(4, 5, 6, 7, 82);
        Assert.AreEqual("101:06:07,082", v.ToString());
    }

    [Test]
    public void TryConvert_Blank()
    {
        XlsxCellValue value = Blank.Value;
        Assert.True(value.TryConvert(out Blank blank));
        Assert.AreEqual(Blank.Value, blank);

        value = String.Empty;
        Assert.True(value.TryConvert(out blank));
        Assert.AreEqual(Blank.Value, blank);
    }

    [Test]
    public void TryConvert_Boolean()
    {
        XlsxCellValue value = true;
        Assert.True(value.TryConvert(out Boolean boolean));
        Assert.True(boolean);

        value = "True";
        Assert.True(value.TryConvert(out boolean));
        Assert.True(boolean);

        value = "False";
        Assert.True(value.TryConvert(out boolean));
        Assert.False(boolean);

        value = 0;
        Assert.True(value.TryConvert(out boolean));
        Assert.False(boolean);

        value = 0.001;
        Assert.True(value.TryConvert(out boolean));
        Assert.True(boolean);
    }

    [Test]
    public void TryConvert_Number()
    {
        var c = CultureInfo.GetCultureInfo("cs-CZ");
        XlsxCellValue value = 5;
        Assert.True(value.TryConvert(out Double number, c));
        Assert.AreEqual(5.0, number);

        value = "1,5";
        Assert.True(value.TryConvert(out number, c));
        Assert.AreEqual(1.5, number);

        value = "1 1/4";
        Assert.True(value.TryConvert(out number, c));
        Assert.AreEqual(1.25, number);

        value = "3.1.1900";
        Assert.True(value.TryConvert(out number, c));
        Assert.AreEqual(3, number);

        value = true;
        Assert.True(value.TryConvert(out number, c));
        Assert.AreEqual(1.0, number);

        value = false;
        Assert.True(value.TryConvert(out number, c));
        Assert.AreEqual(0.0, number);

        value = new DateTime(2020, 4, 5, 10, 14, 5);
        Assert.True(value.TryConvert(out number, c));
        Assert.AreEqual(43926.42644675926, number);

        value = new TimeSpan(18, 0, 0);
        Assert.True(value.TryConvert(out number, c));
        Assert.AreEqual(0.75, number);
    }

    [Test]
    public void TryConvert_DateTime()
    {
        XlsxCellValue v = new DateTime(2020, 1, 1);
        Assert.True(v.TryConvert(out DateTime dt));
        Assert.AreEqual(new DateTime(2020, 1, 1), dt);

        var lastSerialDate = 2958465;
        v = lastSerialDate;
        Assert.True(v.TryConvert(out dt));
        Assert.AreEqual(new DateTime(9999, 12, 31), dt);

        v = lastSerialDate + 1;
        Assert.False(v.TryConvert(out dt));

        v = new TimeSpan(14, 0, 0, 0);
        Assert.True(v.TryConvert(out dt));
        Assert.AreEqual(new DateTime(1900, 1, 14), dt);
    }

    [Test]
    public void TryConvert_TimeSpan()
    {
        var c = CultureInfo.GetCultureInfo("cs-CZ");
        XlsxCellValue v = new TimeSpan(10, 15, 30);
        Assert.True(v.TryConvert(out TimeSpan ts, c));
        Assert.AreEqual(new TimeSpan(10, 15, 30), ts);

        v = "26:15:30,5";
        Assert.True(v.TryConvert(out ts, c));
        Assert.AreEqual(new TimeSpan(1, 2, 15, 30, 500), ts);

        v = 0.75;
        Assert.True(v.TryConvert(out ts, c));
        Assert.AreEqual(new TimeSpan(18, 0, 0), ts);
    }

    [TestCase(1)]
    [TestCase(10)] // microsecond
    [TestCase(3000000001)] // 5 min 1 tick
    public void TimeSpan_can_have_sub_millisecond_precision(long ticks)
    {
        var subMsTimeSpan = TimeSpan.FromTicks(ticks);
        XlsxCellValue value = subMsTimeSpan;
        Assert.AreEqual(subMsTimeSpan, value.GetTimeSpan());
    }

    [TestCase(1)]
    [TestCase(10)] // microsecond
    [TestCase(3000000001)] // 5 min 1 tick
    public void TimeSpan_with_sub_millisecond_precision_is_written_and_loaded_correctly(long ticks)
    {
        // NetFx converts double to string using G15. Core changed it to G17, but Aid.Xlsx still use G15.
        var subMsTimeSpan = TimeSpan.FromTicks(ticks);
        TestHelper.CreateSaveLoadAssert(
            (_, ws) =>
            {
                ws.Cell("A1").Value = subMsTimeSpan;
            },
            (_, ws) =>
            {
                var cellValue = ws.Cell("A1").CachedValue;
                Assert.AreEqual(subMsTimeSpan, cellValue.GetTimeSpan());
            });
    }

    [TestCase(long.MaxValue / (double)TimeSpan.TicksPerDay + 0.01)]
    [TestCase(long.MinValue / (double)TimeSpan.TicksPerDay - 0.01)]
    public void TimeSpan_throws_when_not_representable(double serialDateTime)
    {
        var value = XlsxCellValue.FromSerialTimeSpan(serialDateTime);
        var ex = Assert.Throws<OverflowException>(() => value.GetTimeSpan())!;
        Assert.AreEqual("The serial date time value is too large to be represented in a TimeSpan.", ex.Message);
    }
}
