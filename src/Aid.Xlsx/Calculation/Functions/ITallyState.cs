namespace Aid.Xlsx.Calculation.Functions;

internal interface ITallyState<out TState>
{
    TState Tally(double number);
}
