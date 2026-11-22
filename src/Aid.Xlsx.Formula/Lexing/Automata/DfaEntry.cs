namespace Aid.Xlsx.Formulas.Lexing.Automata;

internal struct DfaEntry
{
    public DfaTransitionEntry[] Transitions;
    public int AcceptSymbolId;
    public DfaEntry(DfaTransitionEntry[] transitions, int acceptSymbolId)
    {
        Transitions = transitions;
        AcceptSymbolId = acceptSymbolId;
    }
}
