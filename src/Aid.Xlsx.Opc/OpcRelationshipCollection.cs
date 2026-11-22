using System.Collections;

namespace Aid.Xlsx.Opc;

/// <summary>Relationships with IDs unique within their source package or part.</summary>
public sealed class OpcRelationshipCollection : IReadOnlyCollection<OpcRelationship>
{
    private readonly Dictionary<string, OpcRelationship> _relationships = new(StringComparer.Ordinal);
    private readonly string? _source;

    internal OpcRelationshipCollection(string? source)
    {
        _source = source;
    }

    public int Count => _relationships.Count;

    public OpcRelationship this[string id] => _relationships[id];

    public OpcRelationship Add(string id, string type, string target, bool isExternal = false)
    {
        var relationship = new OpcRelationship(_source, id, type, target, isExternal);
        if (!_relationships.TryAdd(id, relationship))
            throw new InvalidDataException($"Duplicate OPC relationship ID '{id}'.");
        return relationship;
    }

    public bool Remove(string id) => _relationships.Remove(id);

    public IEnumerator<OpcRelationship> GetEnumerator() => _relationships.Values.GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
