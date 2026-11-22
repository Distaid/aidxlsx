using Aid.Xlsx.Opc;

namespace Aid.Xlsx.Internal.IO;

internal sealed class WorkbookPackage(OpcPackage package)
{
    private const string TransitionalRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
    private const string StrictRelationships = "http://purl.oclc.org/ooxml/officeDocument/relationships/";

    internal OpcPart? WorkbookPart => RelatedParts(package.Relationships, "officeDocument").FirstOrDefault().Part;

    internal static bool IsOfficeRelationship(OpcRelationship relationship, string type) =>
        HasType(relationship.Type, TransitionalRelationships, type) || HasType(relationship.Type, StrictRelationships, type);

    internal (string Id, OpcPart Part) FirstRelatedPart(OpcPart source, string type) =>
        RelatedParts(source, type).FirstOrDefault();

    internal IEnumerable<(string Id, OpcPart Part)> RelatedParts(OpcPart source, string type) =>
        RelatedParts(source.Relationships, type);

    internal OpcPart? RelatedPart(OpcPart source, string id, string type) =>
        source.Relationships.TryGet(id, out var relationship) && IsOfficeRelationship(relationship, type)
            ? Resolve(relationship)
            : null;

    private IEnumerable<(string Id, OpcPart Part)> RelatedParts(OpcRelationshipCollection relationships, string type)
    {
        foreach (var relationship in relationships)
        {
            if (IsOfficeRelationship(relationship, type) && Resolve(relationship) is { } target)
                yield return (relationship.Id, target);
        }
    }

    private OpcPart? Resolve(OpcRelationship relationship) =>
        relationship.TargetPartName is { } name && package.TryGetPart(name, out var part) ? part : null;

    private static bool HasType(string qualifiedType, string prefix, string type) =>
        qualifiedType.StartsWith(prefix, StringComparison.Ordinal) && qualifiedType.AsSpan(prefix.Length).SequenceEqual(type);
}
