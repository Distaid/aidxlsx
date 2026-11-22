using System.IO.Packaging;
using System.Xml.Linq;
using Aid.Xlsx.Opc;

namespace Aid.Xlsx.Internal.IO;

/// <summary>Exposes the OPC container to part readers and writers that still use SDK elements.</summary>
internal sealed class OpcSdkPackage : Package
{
    private const string RelationshipContentType = "application/vnd.openxmlformats-package.relationships+xml";
    private static readonly XNamespace RelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private readonly OpcPackage _opc;
    private readonly Dictionary<string, Part> _parts = new(StringComparer.OrdinalIgnoreCase);

    internal OpcSdkPackage(OpcPackage opc, FileAccess access)
        : base(access)
    {
        _opc = opc;
        AddRelationshipPart(new Uri("/_rels/.rels", UriKind.Relative), opc.Relationships);
        foreach (var part in opc.Parts.Values)
        {
            var uri = new Uri(part.PartName, UriKind.Relative);
            _parts.Add(part.PartName, new Part(this, uri, part.ContentType, part));
            AddRelationshipPart(PackUriHelper.GetRelationshipPartUri(uri), part.Relationships);
        }

        // The base implementation indexes parts on the first GetParts call.
        GetParts();
    }

    internal void WriteTo(Stream destination)
    {
        Flush();
        UpdateRelationships(_opc.Relationships, GetRelationships());
        foreach (var part in _parts.Values)
        {
            if (part.BackingPart is not { } backingPart)
                continue;

            part.StoreChanges();
            UpdateRelationships(backingPart.Relationships, part.GetRelationships());
        }

        _opc.Save(destination);
    }

    protected override PackagePart CreatePartCore(Uri partUri, string contentType, CompressionOption compressionOption)
    {
        var name = partUri.OriginalString;
        var backingPart = PackUriHelper.IsRelationshipPartUri(partUri) ? null : _opc.AddPart(name, contentType);
        var part = new Part(this, partUri, contentType, backingPart, compressionOption);
        _parts.Add(name, part);
        return part;
    }

    protected override PackagePart? GetPartCore(Uri partUri) => _parts.GetValueOrDefault(partUri.OriginalString);

    protected override PackagePart[] GetPartsCore() => _parts.Values.ToArray();

    protected override void DeletePartCore(Uri partUri)
    {
        if (!_parts.Remove(partUri.OriginalString, out var part))
            return;

        part.ReleaseBuffer();
        if (part.BackingPart is { } backingPart)
            _opc.RemovePart(backingPart.PartName);
    }

    protected override void FlushCore()
    {
        // The logical parts are flushed by Package.Flush. The caller chooses the ZIP destination.
    }

    protected override void Dispose(bool disposing)
    {
        try
        {
            base.Dispose(disposing);
        }
        finally
        {
            if (disposing)
            {
                foreach (var part in _parts.Values)
                    part.ReleaseBuffer();
                _parts.Clear();
            }
        }
    }

    private void AddRelationshipPart(Uri uri, OpcRelationshipCollection relationships)
    {
        if (relationships.Count == 0)
            return;

        _parts.Add(uri.OriginalString, new Part(this, uri, RelationshipContentType, null) { InitialRelationships = relationships });
    }

    private static void WriteRelationships(Stream buffer, OpcRelationshipCollection relationships)
    {
        var root = new XElement(RelationshipNamespace + "Relationships");
        foreach (var relationship in relationships)
        {
            var element = new XElement(RelationshipNamespace + "Relationship",
                new XAttribute("Id", relationship.Id), new XAttribute("Type", relationship.Type),
                new XAttribute("Target", relationship.Target));
            if (relationship.IsExternal)
                element.Add(new XAttribute("TargetMode", "External"));
            root.Add(element);
        }

        new XDocument(root).Save(buffer, System.Xml.Linq.SaveOptions.DisableFormatting);
    }

    private static void UpdateRelationships(OpcRelationshipCollection target, IEnumerable<PackageRelationship> source)
    {
        foreach (var existing in target.ToArray())
            target.Remove(existing.Id);

        foreach (var relationship in source)
            target.Add(relationship.Id, relationship.RelationshipType, relationship.TargetUri.OriginalString,
                relationship.TargetMode == TargetMode.External);
    }

    private sealed class Part(OpcSdkPackage package, Uri uri, string contentType, OpcPart? backingPart,
        CompressionOption compression = CompressionOption.Normal) : PackagePart(package, uri, contentType, compression)
    {
        private bool _changed;

        internal OpcPart? BackingPart { get; } = backingPart;

        internal OpcRelationshipCollection? InitialRelationships { get; init; }

        private MemoryStream? Buffer { get; set; }

        protected override Stream GetStreamCore(FileMode mode, FileAccess access)
        {
            if (mode is FileMode.Create or FileMode.CreateNew or FileMode.Truncate)
            {
                Buffer ??= new MemoryStream();
                Buffer.SetLength(0);
                _changed = true;
            }
            else if (Buffer is null)
            {
                Buffer = new MemoryStream();
                if (BackingPart is not null)
                {
                    using var source = BackingPart.OpenRead();
                    source.CopyTo(Buffer);
                }
                else if (InitialRelationships is not null)
                {
                    WriteRelationships(Buffer, InitialRelationships);
                }
            }

            return new PartStream(Buffer, access, mode == FileMode.Append ? Buffer.Length : 0, () => _changed = true);
        }

        internal void StoreChanges()
        {
            if (_changed && Buffer is not null)
                BackingPart!.ReplaceContent(Buffer.GetBuffer().AsSpan(0, (int)Buffer.Length));
        }

        internal void ReleaseBuffer() => Buffer?.Dispose();
    }

    private sealed class PartStream(MemoryStream buffer, FileAccess access, long position, Action onWrite) : Stream
    {
        private bool _closed;

        public override bool CanRead => !_closed && (access & FileAccess.Read) != 0;
        public override bool CanSeek => !_closed;
        public override bool CanWrite => !_closed && (access & FileAccess.Write) != 0;
        public override long Length => buffer.Length;
        public override long Position
        {
            get => position;
            set
            {
                ArgumentOutOfRangeException.ThrowIfNegative(value);
                position = value;
            }
        }

        public override void Flush() => ObjectDisposedException.ThrowIf(_closed, this);

        public override int Read(byte[] bytes, int offset, int count) => Read(bytes.AsSpan(offset, count));

        public override int Read(Span<byte> bytes)
        {
            if (!CanRead)
                throw new NotSupportedException();

            buffer.Position = position;
            var read = buffer.Read(bytes);
            position += read;
            return read;
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            Position = checked(offset + (origin switch
            {
                SeekOrigin.Begin => 0,
                SeekOrigin.Current => position,
                SeekOrigin.End => Length,
                _ => throw new ArgumentOutOfRangeException(nameof(origin)),
            }));
            return position;
        }

        public override void SetLength(long value)
        {
            if (!CanWrite)
                throw new NotSupportedException();

            buffer.SetLength(value);
            onWrite();
        }

        public override void Write(byte[] bytes, int offset, int count) => Write(bytes.AsSpan(offset, count));

        public override void Write(ReadOnlySpan<byte> bytes)
        {
            if (!CanWrite)
                throw new NotSupportedException();

            buffer.Position = position;
            buffer.Write(bytes);
            position += bytes.Length;
            onWrite();
        }

        protected override void Dispose(bool disposing)
        {
            _closed = true;
            base.Dispose(disposing);
        }
    }
}
