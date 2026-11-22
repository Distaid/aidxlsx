using System.Collections.Generic;

namespace Aid.Xlsx.CodeGen.Model.SimpleTypes;

public class SimpleTypeUnion : ISimpleType
{
    public required string Name { get; init; }

    public required List<Restriction> RestrictionsUnion { get; init; } = [];
}
