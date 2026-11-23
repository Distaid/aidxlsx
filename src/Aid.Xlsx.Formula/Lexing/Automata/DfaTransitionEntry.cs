namespace Aid.Xlsx.Formulas.Lexing.Automata;

internal struct DfaTransitionEntry(int[] packedRanges, int destination)
{
    public int[] PackedRanges = packedRanges;
    public int Destination = destination;
}
