namespace Aid.Xlsx.Formulas.Lexing.Automata;

internal struct DfaEntry(DfaTransitionEntry[] transitions, int acceptSymbolId)
{
    public DfaTransitionEntry[] Transitions = transitions;
    public int AcceptSymbolId = acceptSymbolId;
}
