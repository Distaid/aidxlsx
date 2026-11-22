using System.IO.Compression;

namespace Aid.Xlsx.Opc;

/// <summary>A named component of an OPC container.</summary>
public sealed class OpcPart
{
    private readonly OpcPackage _package;
    private ZipArchiveEntry? _entry;
    private byte[] _contents = [];
    private int _contentLength;

    internal OpcPart(OpcPackage package, string partName, string contentType, ZipArchiveEntry? entry = null)
    {
        _package = package;
        _entry = entry;
        PartName = partName;
        ContentType = contentType;
        Relationships = new OpcRelationshipCollection(partName);
    }

    public string PartName { get; }

    public string ContentType { get; private set; }

    public long Length => _entry?.Length ?? _contentLength;

    public OpcRelationshipCollection Relationships { get; }

    /// <summary>Opens a stream over the current contents. The caller disposes the stream.</summary>
    public Stream OpenRead()
    {
        _package.ThrowIfDisposed();
        return _entry?.Open() ?? new MemoryStream(_contents, 0, _contentLength, writable: false);
    }

    /// <summary>Opens a stream that replaces the part contents when disposed.</summary>
    public Stream OpenWrite()
    {
        _package.ThrowIfDisposed();
        return new PartWriteStream(this);
    }

    public void ChangeContentType(string contentType)
    {
        _package.ThrowIfDisposed();
        OpcMetadata.ValidateContentType(contentType);
        if (contentType.Equals(OpcMetadata.RelationshipsContentType, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The relationships content type is reserved.");
        ContentType = contentType;
    }

    /// <summary>Replaces the contents with a copy of the supplied data.</summary>
    public void ReplaceContent(ReadOnlySpan<byte> contents)
    {
        _package.ThrowIfDisposed();
        SetContent(contents.ToArray(), contents.Length);
    }

    private void SetContent(byte[] contents, int length)
    {
        _contents = contents;
        _contentLength = length;
        _entry = null;
    }

    private sealed class PartWriteStream(OpcPart part) : MemoryStream
    {
        protected override void Dispose(bool disposing)
        {
            if (disposing && CanWrite)
            {
                part._package.ThrowIfDisposed();
                part.SetContent(GetBuffer(), checked((int)Length));
            }

            base.Dispose(disposing);
        }
    }
}
