using System.Collections.Generic;

namespace Aid.Xlsx.CodeGen.Schema.Model.Elements;

/// <summary>
/// <c><![CDATA[<xsd:sequence>]]></c> inside <c><![CDATA[<xsd:complexType>]]></c>.
/// </summary>
public class Sequence : IElementGroup
{
    public required List<IElementGroup> Children { get; init; } = [];

    public required Occurrences Occurrences { get; init; }
}
