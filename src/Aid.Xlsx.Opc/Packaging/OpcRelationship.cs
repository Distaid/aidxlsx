using System.Xml;

namespace Aid.Xlsx.Opc.Packaging;

/// <summary>A package or part relationship. The original URI reference is preserved.</summary>
public sealed class OpcRelationship
{
    internal OpcRelationship(string? source, string id, string type, string target, bool isExternal)
    {
        XmlConvert.VerifyNCName(id);
        if (!Uri.TryCreate(type, UriKind.Absolute, out _))
            throw new InvalidDataException($"Invalid OPC relationship type '{type}'.");

        ArgumentException.ThrowIfNullOrEmpty(target);
        if (isExternal)
        {
            var hasControlCharacters = target.Any(char.IsControl);
            if (hasControlCharacters || !Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out _))
                throw new InvalidDataException($"Invalid OPC relationship target '{target}'.");
        }

        Id = id;
        Type = type;
        Target = target;
        IsExternal = isExternal;
        TargetPartName = isExternal ? null : OpcPartName.ResolveTarget(source, target);
    }

    public string Id { get; }

    public string Type { get; }

    public string Target { get; }

    public bool IsExternal { get; }

    /// <summary>The absolute package part name, without the optional fragment, for an internal target.</summary>
    public string? TargetPartName { get; }
}
