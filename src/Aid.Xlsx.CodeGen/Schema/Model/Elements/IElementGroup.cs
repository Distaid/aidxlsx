using System.Collections.Generic;

namespace Aid.Xlsx.CodeGen.Schema.Model.Elements;

/// <summary>
/// A node in a complex type element tree.
/// </summary>
public interface IElementGroup
{
    /// <summary>
    /// Children elements.
    /// </summary>
    public List<IElementGroup> Children { get; }
}
