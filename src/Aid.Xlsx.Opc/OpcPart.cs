using System.IO.Compression;

namespace Aid.Xlsx.Opc;

/// <summary>A named component of an OPC container.</summary>
public sealed class OpcPart
{
    private readonly OpcPackage _package;
    private ZipArchiveEntry? _entry;
    private byte[] _contents = [];

    internal OpcPart(OpcPackage package, string partName, string contentType, ZipArchiveEntry? entry = null)
    {
        _package = package;
        _entry = entry;
        PartName = partName;
        ContentType = contentType;
        Relationships = new OpcRelationshipCollection(partName);
    }

    public string PartName { get; }

    public string ContentType { get; }

    public OpcRelationshipCollection Relationships { get; }

    /// <summary>Opens a stream over the current contents. The caller disposes the stream.</summary>
    public Stream OpenRead()
    {
        _package.ThrowIfDisposed();
        return _entry?.Open() ?? new MemoryStream(_contents, writable: false);
    }

    /// <summary>Replaces the contents with a copy of the supplied data.</summary>
    public void ReplaceContent(ReadOnlySpan<byte> contents)
    {
        _package.ThrowIfDisposed();
        _contents = contents.ToArray();
        _entry = null;
    }
}
