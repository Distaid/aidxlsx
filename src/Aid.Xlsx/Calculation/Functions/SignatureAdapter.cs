#nullable disable

namespace Aid.Xlsx.Calculation.Functions;

/// <summary>
/// A collection of adapter functions from a more a generic formula function to more specific ones.
/// </summary>
internal static class SignatureAdapter
{
    #region Signature adapters
    // Each method converts a more specific signature of a function into a generic formula function type.
    // We have many functions with same signature and the adapters should be reusable. Convert parameters
    // through value converters below. We can hopefully generate them at a later date, so try to keep them similar.

    public static CalcEngineFunction Adapt(Func<ScalarValue> f)
    {
        return (_, _) => f().ToAnyValue();
    }

    public static CalcEngineFunction AdaptCoerced(Func<bool, AnyValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<bool, XlsxError> arg0Converted = CoerceToLogical(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            return f(arg0);
        };
    }

    public static CalcEngineFunction Adapt(Func<double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            return f(arg0).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            return f(ctx, arg0).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, ScalarValue, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<ScalarValue, XlsxError> arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out ScalarValue arg0, out XlsxError err0))
                return err0;

            return f(ctx, arg0).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, double, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, XlsxError> arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, double, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<string, XlsxError> arg2Converted = ToText(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, double, double, bool, AnyValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, XlsxError> arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            OneOf<bool, XlsxError> arg3Converted = CoerceToLogical(args[3], ctx);
            if (!arg3Converted.TryPickT0(out var arg3, out XlsxError err3))
                return err3;

            return f(ctx, arg0, arg1, arg2, arg3);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            return f(ctx, arg0).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<string, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<string, XlsxError> arg1Converted = ToText(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            return f(arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, double, double, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, XlsxError> arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            OneOf<string, XlsxError> arg3Converted = ToText(args[3], ctx);
            if (!arg3Converted.TryPickT0(out var arg3, out XlsxError err3))
                return err3;

            return f(ctx, arg0, arg1, arg2, arg3).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, double, double, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, XlsxError> arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, AnyValue, double, AnyValue> f)
    {
        return (ctx, args) =>
        {
            AnyValue arg0 = args[0];

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            return f(ctx, arg0, arg1);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, ScalarValue?, AnyValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            var arg1 = default(ScalarValue?);
            if (args.Length > 1)
            {
                OneOf<ScalarValue, XlsxError> arg1Converted = ToScalarValue(args[1], ctx);
                if (!arg1Converted.TryPickT0(out ScalarValue arg1Value, out XlsxError err1))
                    return err1;

                arg1 = arg1Value;
            }


            return f(ctx, arg0, arg1);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, List<string>, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var texts = new List<string>(args.Length);
            foreach (AnyValue arg in args)
            {
                OneOf<string, XlsxError> argConverted = ToText(arg, ctx);
                if (!argConverted.TryPickT0(out var text, out XlsxError error))
                    return error;

                texts.Add(text);
            }

            return f(ctx, texts).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, AnyValue, AnyValue> f)
    {
        return (ctx, args) => f(ctx, args[0]);
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, ScalarValue, AnyValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<ScalarValue, XlsxError> arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out ScalarValue arg0, out XlsxError err0))
                return err0;

            return f(ctx, arg0);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, ScalarValue, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<ScalarValue, XlsxError> arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out ScalarValue arg0, out XlsxError err0))
                return err0;

            OneOf<string, XlsxError> arg1Converted = ToText(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<ScalarValue, ScalarValue, AnyValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<ScalarValue, XlsxError> arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out ScalarValue arg0, out XlsxError err0))
                return err0;

            OneOf<ScalarValue, XlsxError> arg1Converted = ToScalarValue(args[1], ctx);
            if (!arg1Converted.TryPickT0(out ScalarValue arg1, out XlsxError err1))
                return err1;

            return f(arg0, arg1);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, AnyValue, ScalarValue, AnyValue> f)
    {
        return (ctx, args) =>
        {
            AnyValue arg0 = args[0];

            OneOf<ScalarValue, XlsxError> arg1Converted = ToScalarValue(args[1], ctx);
            if (!arg1Converted.TryPickT0(out ScalarValue arg1, out XlsxError err1))
                return err1;

            return f(ctx, arg0, arg1);
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, List<Array>, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            var arrays = new List<Array>();
            foreach (AnyValue arg in args)
            {
                if (arg.TryPickSingleOrMultiValue(out ScalarValue scalar, out Array array, ctx))
                    array = new ScalarArray(scalar, 1, 1);

                arrays.Add(array);
            }

            return f(ctx, arrays).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, string, bool, List<AnyValue>, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<bool, XlsxError> arg1Converted = CoerceToLogical(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            var remainingArgs = new List<AnyValue>();
            foreach (AnyValue arg in args[2..])
                remainingArgs.Add(arg);

            return f(ctx, arg0, arg1, remainingArgs).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<ScalarValue, AnyValue, AnyValue, AnyValue> f, AnyValue lastDefault)
    {
        return (ctx, args) =>
        {
            OneOf<ScalarValue, XlsxError> arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out ScalarValue arg0, out XlsxError err0))
                return err0;

            AnyValue arg1 = args[1];
            AnyValue arg2 = args.Length > 2 ? args[2] : lastDefault;
            return f(arg0, arg1, arg2);
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, double, double, ScalarValue> f, double lastDefault)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args.Length > 1 ? args[1] : lastDefault, ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, double, double, double, ScalarValue> f, double lastDefault)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, XlsxError> arg2Converted = ToNumber(args.Length > 2 ? args[2] : lastDefault, ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, double, double, bool, ScalarValue> f, bool lastDefault)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<bool, XlsxError> arg2Converted = CoerceToLogical(args.Length > 2 ? args[2] : lastDefault, ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, string, double, ScalarValue> f, double lastDefault)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args.Length > 1 ? args[1] : lastDefault, ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<double, double, double, ScalarValue> f, double lastDefault)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, XlsxError> arg2Converted = ToNumber(args.Length > 2 ? args[2] : lastDefault, ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction Adapt(Func<CalcContext, double, AnyValue[], AnyValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            AnyValue[] argsLoop = args[1..].ToArray();
            return f(ctx, arg0, argsLoop);
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, string, string, OneOf<double, Blank>, AnyValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<string, XlsxError> arg1Converted = ToText(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, Blank> arg2Optional = Blank.Value;
            if (args.Length > 2)
            {
                OneOf<double, XlsxError> arg2Converted = ToNumber(args[2], ctx);
                if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                    return err2;

                arg2Optional = arg2;
            }

            return f(ctx, arg0, arg1, arg2Optional);
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, ScalarValue, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<ScalarValue, XlsxError> arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out ScalarValue arg0, out XlsxError err0))
                return err0;

            OneOf<ScalarValue, XlsxError> arg1Converted = args.Length > 1 ? ToScalarValue(args[1], ctx) : ScalarValue.Blank;
            if (!arg1Converted.TryPickT0(out ScalarValue arg1, out XlsxError err1))
                return err1;

            return f(ctx, arg0, arg1).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, ScalarValue, AnyValue, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<ScalarValue, XlsxError> arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out ScalarValue arg0, out XlsxError err0))
                return err0;

            OneOf<ScalarValue, XlsxError> arg1Converted = ToScalarValue(args[1], ctx);
            if (!arg1Converted.TryPickT0(out ScalarValue arg1, out XlsxError err1))
                return err1;

            AnyValue arg2 = args.Length > 2 ? args[2] : AnyValue.Blank;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, AnyValue, ScalarValue, AnyValue, AnyValue> f)
    {
        return (ctx, args) =>
        {
            AnyValue arg0 = args[0];

            OneOf<ScalarValue, XlsxError> arg1Converted = ToScalarValue(args[1], ctx);
            if (!arg1Converted.TryPickT0(out ScalarValue arg1, out XlsxError err1))
                return err1;

            AnyValue arg2 = args.Length > 2 ? args[2] : AnyValue.Blank;

            return f(ctx, arg0, arg1, arg2);
        };
    }

    /// <summary>
    /// An adapter for <c>{SUM,AVERAGE}IFS</c> functions.
    /// </summary>
    public static CalcEngineFunction AdaptIfs(Func<CalcContext, AnyValue, List<(AnyValue Range, ScalarValue Criteria)>, AnyValue> f)
    {
        return (ctx, args) =>
        {
            AnyValue tallyRange = args[0];
            if (!ToCriteria(ctx, args[1..]).TryPickT0(out List<(AnyValue Range, ScalarValue Criteria)> criteria, out XlsxError error))
                return error;

            return f(ctx, tallyRange, criteria);
        };
    }

    /// <summary>
    /// An adapter for <c>COUNTIFS</c> function.
    /// </summary>
    public static CalcEngineFunction AdaptIfs(Func<CalcContext, List<(AnyValue Range, ScalarValue Criteria)>, AnyValue> f)
    {
        return (ctx, args) =>
        {
            if (!ToCriteria(ctx, args).TryPickT0(out List<(AnyValue Range, ScalarValue Criteria)> criteria, out XlsxError error))
                return error;

            return f(ctx, criteria);
        };
    }

    public static CalcEngineFunction AdaptIndex(Func<CalcContext, AnyValue, List<int>, AnyValue> f)
    {
        return (ctx, args) =>
        {
            AnyValue arg0 = args[0];
            var numbers = new List<int>(args.Length - 1);
            for (var i = 1; i < args.Length; ++i)
            {
                if (!ToNumber(args[i], ctx).TryPickT0(out var number, out XlsxError error))
                    return error;

                numbers.Add((int)number);
            }

            return f(ctx, arg0, numbers);
        };
    }

    public static CalcEngineFunction AdaptMatch(Func<CalcContext, ScalarValue, AnyValue, int, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<ScalarValue, XlsxError> arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out ScalarValue arg0, out XlsxError err0))
                return err0;

            AnyValue arg1 = args[1];
            OneOf<double, XlsxError> arg2Converted = args.Length > 2 ? ToNumber(args[2], ctx) : 1;
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(ctx, arg0, arg1, (int)arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptSeriesSum(Func<CalcContext, double, double, double, Array, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            // SERIESSUM doesn't convert logical values to number...
            if (args[0].IsLogical)
                return XlsxError.IncompatibleValue;

            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            if (args[1].IsLogical)
                return XlsxError.IncompatibleValue;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            if (args[2].IsLogical)
                return XlsxError.IncompatibleValue;

            OneOf<double, XlsxError> arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            if (args[3].TryPickSingleOrMultiValue(out ScalarValue scalar, out Array arg3, ctx))
            {
                if (scalar.IsLogical)
                    return XlsxError.IncompatibleValue;

                if (!scalar.ToNumber(ctx.Culture).TryPickT0(out var number, out XlsxError error))
                    return error;

                arg3 = new ScalarArray(number, 1, 1);
            }

            return f(ctx, arg0, arg1, arg2, arg3).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptNumberValue(Func<CalcContext, string, string, string, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            var decimalSeparator = ctx.Culture.NumberFormat.NumberDecimalSeparator;
            OneOf<string, XlsxError> arg1Converted = ToText(args.Length > 1 ? args[1] : decimalSeparator, ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            var groupSeparator = ctx.Culture.NumberFormat.NumberGroupSeparator;
            OneOf<string, XlsxError> arg2Converted = ToText(args.Length > 2 ? args[2] : groupSeparator, ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptSubstitute(Func<CalcContext, string, string, string, double?, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            OneOf<string, XlsxError> arg0Converted = ToText(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<string, XlsxError> arg1Converted = ToText(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<string, XlsxError> arg2Converted = ToText(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            double? arg3 = null;
            if (args.Length > 3)
            {
                // Excel doesn't accept logical, be more permissive.
                OneOf<double, XlsxError> arg3Converted = ToNumber(args[3], ctx);
                if (!arg3Converted.TryPickT0(out var arg3Number, out XlsxError err3))
                    return err3;

                arg3 = arg3Number;
            }

            return f(ctx, arg0, arg1, arg2, arg3).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptMultinomial(Func<CalcContext, List<IEnumerable<ScalarValue>>, ScalarValue> f)
    {
        return (ctx, args) =>
        {
            // This can skip blank values, because blank doesn't increase nominator
            // and doesn't change denominator due to 0! = 1
            var scalarCollections = new List<IEnumerable<ScalarValue>>(args.Length);
            foreach (AnyValue arg in args)
                scalarCollections.Add(GetNonBlankScalars(arg, ctx));

            return f(ctx, scalarCollections).ToAnyValue();
        };

        static IEnumerable<ScalarValue> GetNonBlankScalars(AnyValue value, CalcContext ctx)
        {
            if (value.TryPickScalar(out ScalarValue scalar, out OneOf<Array, Reference> collection))
            {
                if (!scalar.IsBlank)
                    yield return scalar;
            }
            else if (collection.TryPickT0(out Array array, out Reference reference))
            {
                foreach (ScalarValue element in array)
                {
                    if (!element.IsBlank)
                        yield return element;
                }
            }
            else
            {
                foreach (ScalarValue element in ctx.GetNonBlankValues(reference))
                {
                    if (!element.IsBlank)
                        yield return element;
                }
            }
        }
    }

    /// <summary>
    /// Adapt a function that accepts areas as arguments (e.g. SUMPRODUCT). The key benefit is
    /// that all <c>ReferenceArray</c> allocation is done once for a function. The method
    /// shouldn't be used for functions that accept 3D references (e.g. SUMSQ). It is still
    /// necessary to check all errors in the <paramref name="f"/>, adapt method doesn't do that
    /// on its own (potential performance problem). The signature uses an array instead of
    /// IReadOnlyList interface for performance reasons (can't JIT access props through interface).
    /// </summary>
    public static CalcEngineFunction Adapt(Func<CalcContext, Array[], AnyValue> f)
    {
        return (ctx, args) =>
        {
            var areas = new Array[args.Length];
            for (var i = 0; i < args.Length; ++i)
            {
                areas[i] = args[i].TryPickSingleOrMultiValue(out ScalarValue scalar, out Array array, ctx)
                    ? new ScalarArray(scalar, 1, 1)
                    : array;
            }

            return f(ctx, areas);
        };
    }

    public static CalcEngineFunction AdaptLastOptional(Func<CalcContext, ScalarValue, AnyValue, double, bool, AnyValue> f, bool defaultValue0)
    {
        return (ctx, args) =>
        {
            OneOf<ScalarValue, XlsxError> arg0Converted = ToScalarValue(args[0], ctx);
            if (!arg0Converted.TryPickT0(out ScalarValue arg0, out XlsxError err0))
                return err0;

            AnyValue arg1 = args[1];

            OneOf<double, XlsxError> arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            OneOf<bool, XlsxError> arg3Converted = args.Length >= 4 ? CoerceToLogical(args[3], ctx) : defaultValue0;
            if (!arg3Converted.TryPickT0(out var arg3, out XlsxError err3))
                return err3;

            return f(ctx, arg0, arg1, arg2, arg3);
        };
    }

    public static CalcEngineFunction AdaptLastTwoOptional(Func<double, double, double, ScalarValue> f, double defaultValue1, double defaultValue2)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = args.Length > 1 ? ToNumber(args[1], ctx) : defaultValue1;
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, XlsxError> arg2Converted = args.Length > 2 ? ToNumber(args[2], ctx) : defaultValue2;
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastTwoOptional(Func<CalcContext, double, double, bool, ScalarValue> f, double defaultValue1, bool defaultValue2)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = args.Length > 1 ? ToNumber(args[1], ctx) : defaultValue1;
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            // AnyValue to bool has different semantic than AnyValue to number, e.g. "0" is not valid for bool coercion
            AnyValue arg2Converted = args.Length > 2 ? args[2] : defaultValue2;
            if (!CoerceToLogical(arg2Converted, ctx).TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            return f(ctx, arg0, arg1, arg2).ToAnyValue();
        };
    }

    public static CalcEngineFunction AdaptLastTwoOptional(Func<double, double, double, double, double, AnyValue> f, double defaultValue0, double defaultValue1)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, XlsxError> arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            var arg3Optional = defaultValue0;
            if (args.Length >= 4)
            {
                OneOf<double, XlsxError> arg3Converted = ToNumber(args[3], ctx);
                if (!arg3Converted.TryPickT0(out var arg3, out XlsxError err3))
                    return err3;

                arg3Optional = arg3;
            }

            var arg4Optional = defaultValue1;
            if (args.Length >= 5)
            {
                OneOf<double, XlsxError> arg4Converted = ToNumber(args[4], ctx);
                if (!arg4Converted.TryPickT0(out var arg4, out XlsxError err4))
                    return err4;

                arg4Optional = arg4;
            }

            return f(arg0, arg1, arg2, arg3Optional, arg4Optional);
        };
    }

    public static CalcEngineFunction AdaptLastTwoOptional(Func<double, double, double, double, double, double, AnyValue> f, double defaultValue0, double defaultValue1)
    {
        return (ctx, args) =>
        {
            OneOf<double, XlsxError> arg0Converted = ToNumber(args[0], ctx);
            if (!arg0Converted.TryPickT0(out var arg0, out XlsxError err0))
                return err0;

            OneOf<double, XlsxError> arg1Converted = ToNumber(args[1], ctx);
            if (!arg1Converted.TryPickT0(out var arg1, out XlsxError err1))
                return err1;

            OneOf<double, XlsxError> arg2Converted = ToNumber(args[2], ctx);
            if (!arg2Converted.TryPickT0(out var arg2, out XlsxError err2))
                return err2;

            OneOf<double, XlsxError> arg3Converted = ToNumber(args[3], ctx);
            if (!arg3Converted.TryPickT0(out var arg3, out XlsxError err3))
                return err3;

            var arg4Optional = defaultValue0;
            if (args.Length >= 5)
            {
                OneOf<double, XlsxError> arg4Converted = ToNumber(args[4], ctx);
                if (!arg4Converted.TryPickT0(out var arg4, out XlsxError err4))
                    return err4;

                arg4Optional = arg4;
            }

            var arg5Optional = defaultValue1;
            if (args.Length >= 6)
            {
                OneOf<double, XlsxError> arg5Converted = ToNumber(args[5], ctx);
                if (!arg5Converted.TryPickT0(out var arg5, out XlsxError err5))
                    return err5;

                arg5Optional = arg5;
            }

            return f(arg0, arg1, arg2, arg3, arg4Optional, arg5Optional);
        };
    }

    #endregion

    #region Value converters
    // Each method is named ToSomething and it converts an argument into a desired type (e.g. for ToSomething it should be type Something).
    // Return value is always OneOf<Something, Error>, if there is an error, return it as an error.

    private static OneOf<bool, XlsxError> CoerceToLogical(in AnyValue value, CalcContext ctx)
    {
        if (!ToScalarValue(in value, ctx).TryPickT0(out ScalarValue scalar, out XlsxError scalarError))
            return scalarError;

        // LibreOffice does accept text, tries to parse it as a number and coerces the number
        // to bool. Excel does not accept number in text argument.
        if (!scalar.TryCoerceLogicalOrBlankOrNumberOrText(out var logical, out XlsxError coercionError))
            return coercionError;

        return logical;
    }

    private static OneOf<double, XlsxError> ToNumber(in AnyValue value, CalcContext ctx)
    {
        if (value.TryPickScalar(out ScalarValue scalar, out OneOf<Array, Reference> collection))
            return scalar.ToNumber(ctx.Culture);

        // When user specifies array as an argument in an array formula for a scalar function, use [0,0]
        if (collection.TryPickT0(out Array array, out Reference reference))
            return array[0, 0].ToNumber(ctx.Culture);

        if (reference.TryGetSingleCellValue(out ScalarValue scalarValue, ctx))
            return scalarValue.ToNumber(ctx.Culture);

        throw new NotImplementedException("Array formulas not implemented.");
    }

    private static OneOf<string, XlsxError> ToText(in AnyValue value, CalcContext ctx)
    {
        if (value.TryPickScalar(out ScalarValue scalar, out OneOf<Array, Reference> collection))
            return scalar.ToText(ctx.Culture);

        if (collection.TryPickT0(out _, out Reference reference))
            throw new NotImplementedException("Array formulas not implemented.");

        if (reference.TryGetSingleCellValue(out ScalarValue scalarValue, ctx))
            return scalarValue.ToText(ctx.Culture);

        throw new NotImplementedException("Array formulas not implemented.");
    }

    private static OneOf<ScalarValue, XlsxError> ToScalarValue(in AnyValue value, CalcContext ctx)
    {
        if (value.TryPickScalar(out ScalarValue scalar, out OneOf<Array, Reference> collection))
            return scalar;

        if (collection.TryPickT0(out Array array, out Reference reference))
            return array[0, 0];

        if (reference.TryGetSingleCellValue(out ScalarValue referenceScalar, ctx))
            return referenceScalar;

        return OneOf<ScalarValue, XlsxError>.FromT1(XlsxError.IncompatibleValue);
    }

    private static OneOf<List<(AnyValue Range, ScalarValue Criteria)>, XlsxError> ToCriteria(CalcContext ctx, ReadOnlySpan<AnyValue> args)
    {
        var allCriteria = new List<(AnyValue Range, ScalarValue Criteria)>();
        var pairCount = (args.Length + 1) / 2;
        for (var i = 0; i < pairCount; ++i)
        {
            var rangeArgIndex = 2 * i;
            AnyValue range = args[rangeArgIndex];

            // Excel grammar requires even number of arguments. We can't
            // do that, so use blank for missing pair value.
            var criteriaArgIndex = rangeArgIndex + 1;
            OneOf<ScalarValue, XlsxError> criteriaArgConverted = criteriaArgIndex < args.Length
                ? ToScalarValue(args[criteriaArgIndex], ctx)
                : ScalarValue.Blank;
            if (!criteriaArgConverted.TryPickT0(out ScalarValue criteria, out XlsxError criteriaError))
                return criteriaError;

            allCriteria.Add((range, criteria));
        }

        return allCriteria;
    }
    #endregion
}
