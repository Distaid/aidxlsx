namespace Aid.Xlsx.Formulas.Rolex;

internal struct DfaTransitionEntry
{
    public int[] PackedRanges;
    public int Destination;

    public DfaTransitionEntry(int[] packedRanges, int destination)
    {
        PackedRanges = packedRanges;
        Destination = destination;
    }
}
