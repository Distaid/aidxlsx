namespace Aid.Xlsx.Calculation;

/// <summary>
/// A non-state representation of a formula that can be used by many cells.
/// </summary>
internal class Formula(string text, ValueNode root)
{

    /// <summary>Text of the formula.</summary>
    public string Text { get; } = text;

    public ValueNode AstRoot { get; } = root;
}
