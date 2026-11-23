namespace Aid.Xlsx.Formulas.References.Transforms;

/// <summary>
/// A range of a symbol in formula text.
/// </summary>
/// <remarks>
/// Create a substring of a symbol.
/// </remarks>
public readonly struct SymbolRange(int startIndex, int endIndex)
{

    /// <summary>
    /// Start index of symbol in formula text.
    /// </summary>
    public int Start { get; } = startIndex;

    /// <summary>
    /// End index of symbol in formula text. Can be outside of text bounds, if symbol ends at the
    /// last char of formula.
    /// </summary>
    public int End { get; } = endIndex;

    /// <summary>
    /// Length of a symbol.
    /// </summary>
    public int Length => End - Start;

    /// <summary>
    /// Get range indexes.
    /// </summary>
    public override string ToString()
    {
        return $"[{Start}:{End}]";
    }
}
