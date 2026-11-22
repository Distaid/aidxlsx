using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Aid.Xlsx.CodeGen.Schema.Model.TopLevel;

namespace Aid.Xlsx.CodeGen.Schema.Model;

/// <summary>
/// A representation of a one XSD file.
/// </summary>
public class XsdSchema
{
    /// <summary>
    /// Imports in the file.
    /// </summary>
    public List<ImportElement> Imports { get; } = [];

    /// <summary>
    /// One of <c>xsd:attributeGroup</c>, <c>xsd:complexType</c>, <c>xsd:element</c>, <c>xsd:group</c> or <c>xsd:simpleType</c>.
    /// </summary>
    public List<object> Entries { get; } = [];

    internal bool TryGetParslet(ParsletName parsletName, [NotNullWhen(true)] out IParslet? parslet)
    {
        parslet = Entries.OfType<IParslet>().SingleOrDefault(x => x.Name == parsletName);
        return parslet is not null;
    }
}
