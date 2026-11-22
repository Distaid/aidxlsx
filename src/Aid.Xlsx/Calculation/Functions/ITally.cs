namespace Aid.Xlsx.Calculation.Functions;

internal interface ITally
{
    OneOf<T, XlsxError> Tally<T>(CalcContext ctx, Span<AnyValue> args, T initialState)
        where T : ITallyState<T>;
}
