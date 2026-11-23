using System.Collections.ObjectModel;
using System.Diagnostics.CodeAnalysis;
using System.IO.Compression;

namespace Aid.Xlsx.Opc.Packaging;

/// <summary>An Open Packaging Conventions container backed by a ZIP archive.</summary>
/// <remarks>Package operations and streams over the original ZIP archive are not thread safe.</remarks>
public sealed class OpcPackage : IDisposable
{
    private static readonly DateTimeOffset ZipEpoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);
    private readonly Dictionary<string, OpcPart> _parts = [with(StringComparer.OrdinalIgnoreCase)];
    private readonly Dictionary<string, string> _defaults = [with(StringComparer.OrdinalIgnoreCase)];
    private readonly Stream? _source;
    private readonly ZipArchive? _archive;
    private bool _disposed;

    private OpcPackage(Stream? source)
    {
        Parts = new ReadOnlyDictionary<string, OpcPart>(_parts);
        Relationships = new OpcRelationshipCollection(null);
        _source = source;
        _archive = source is null ? null : new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
    }

    public IReadOnlyDictionary<string, OpcPart> Parts { get; }

    public OpcRelationshipCollection Relationships { get; }

    /// <summary>Creates an empty container.</summary>
    public static OpcPackage Create() => new(null);

    /// <summary>Opens a container over a seekable stream, leaving it open when the package is disposed.</summary>
    public static OpcPackage Open(Stream source, OpcOpenOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(source);
        if (!source.CanRead || !source.CanSeek)
            throw new ArgumentException("The source stream must be readable and seekable.", nameof(source));

        options ??= new OpcOpenOptions();
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxEntries);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxUncompressedBytes);
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(options.MaxMetadataCharacters);
        if (!double.IsFinite(options.MaxCompressionRatio) || options.MaxCompressionRatio < 1)
            throw new ArgumentOutOfRangeException(nameof(options), "The compression ratio must be finite and at least one.");

        var package = new OpcPackage(source);
        try
        {
            package.Load(package._archive!, options);
            return package;
        }
        catch
        {
            package.Dispose();
            throw;
        }
    }

    public OpcPart GetPart(string partName)
    {
        ThrowIfDisposed();
        return _parts[partName];
    }

    public bool TryGetPart(string partName, [NotNullWhen(true)] out OpcPart? part)
    {
        ThrowIfDisposed();
        return _parts.TryGetValue(partName, out part);
    }

    public OpcPart AddPart(string partName, string contentType)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(partName);
        OpcPartName.Validate(partName);
        OpcMetadata.ValidateContentType(contentType);
        if (OpcPartName.TryGetRelationshipSource(partName[1..], out _)
            || contentType.Equals(OpcMetadata.RelationshipsContentType, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException($"OPC part '{partName}' is reserved for relationships.");

        var part = new OpcPart(this, partName, contentType);
        if (!_parts.TryAdd(partName, part))
            throw new InvalidDataException($"Duplicate OPC part '{partName}'.");

        return part;
    }

    /// <summary>Removes a part and all relationships targeting it.</summary>
    public bool RemovePart(string partName)
    {
        ThrowIfDisposed();
        if (!_parts.Remove(partName))
            return false;

        RemoveLinks(Relationships, partName);
        foreach (var part in _parts.Values)
            RemoveLinks(part.Relationships, partName);
        return true;
    }

    /// <summary>Writes the package to a separate stream, leaving both streams open.</summary>
    public void Save(Stream destination)
    {
        ThrowIfDisposed();
        ArgumentNullException.ThrowIfNull(destination);
        if (!destination.CanWrite || ReferenceEquals(destination, _source))
            throw new ArgumentException("The destination must be a separate writable stream.", nameof(destination));

        var defaults = new Dictionary<string, string>(_defaults, _defaults.Comparer);
        var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var entries = new SortedDictionary<string, Action<Stream>>(StringComparer.Ordinal)
        {
            { OpcPartName.ContentTypesEntry, stream => OpcMetadata.WriteContentTypes(stream, defaults, overrides) }
        };
        AddRelationshipEntry(null, Relationships);
        foreach (var part in _parts.Values)
        {
            DeclareContentType(part.PartName, part.ContentType);
            entries.Add(part.PartName[1..], destinationPart =>
            {
                using var sourcePart = part.OpenRead();
                sourcePart.CopyTo(destinationPart);
            });
            AddRelationshipEntry(part.PartName, part.Relationships);
        }

        if (destination.CanSeek)
        {
            destination.Position = 0;
            destination.SetLength(0);
        }

        using var zip = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
        foreach (var (name, write) in entries)
        {
            var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            entry.LastWriteTime = ZipEpoch;
            entry.ExternalAttributes = 0;
            using var contents = entry.Open();
            write(contents);
        }

        void AddRelationshipEntry(string? source, OpcRelationshipCollection relationships)
        {
            if (relationships.Count == 0)
                return;

            var name = OpcPartName.RelationshipEntry(source);
            defaults.TryAdd("rels", OpcMetadata.RelationshipsContentType);
            DeclareContentType("/" + name, OpcMetadata.RelationshipsContentType);
            entries.Add(name, stream => OpcMetadata.WriteRelationships(stream, relationships));
        }

        void DeclareContentType(string partName, string contentType)
        {
            if (!defaults.TryGetValue(GetExtension(partName), out var defaultType) || defaultType != contentType)
                overrides.Add(partName, contentType);
        }
    }

    public void Dispose()
    {
        _disposed = true;
        _archive?.Dispose();
    }

    internal void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);

    private void Load(ZipArchive archive, OpcOpenOptions options)
    {
        if (archive.Entries.Count > options.MaxEntries)
            throw new InvalidDataException("The OPC container has too many ZIP entries.");

        var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        long totalSize = 0;
        foreach (var entry in archive.Entries)
        {
            if (entry.Length < 0 || entry.Length > options.MaxUncompressedBytes - totalSize)
                throw new InvalidDataException("The OPC container exceeds its uncompressed size limit.");
            totalSize += entry.Length;
            if (entry.Length > 0 && (entry.CompressedLength <= 0 || (double)entry.Length / entry.CompressedLength > options.MaxCompressionRatio))
                throw new InvalidDataException($"OPC entry '{entry.FullName}' exceeds the compression ratio limit.");

            if (!entries.TryAdd(entry.FullName, entry))
                throw new InvalidDataException($"Duplicate OPC entry '{entry.FullName}'.");
        }

        if (!entries.Remove(OpcPartName.ContentTypesEntry, out var contentTypes))
            throw new InvalidDataException("The OPC container is missing [Content_Types].xml.");

        var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        OpcMetadata.ReadContentTypes(contentTypes, options.MaxMetadataCharacters, _defaults, overrides);

        var relationshipEntries = new List<(string? Source, ZipArchiveEntry Entry)>();
        foreach (var (name, entry) in entries)
        {
            if (name.EndsWith('/') && entry.Length == 0)
            {
                OpcPartName.Validate("/" + name[..^1]);
                continue;
            }

            var partName = OpcPartName.Validate("/" + name);
            if (!overrides.TryGetValue(partName, out var contentType) && !_defaults.TryGetValue(GetExtension(partName), out contentType))
                throw new InvalidDataException($"OPC part '{partName}' has no content type.");

            var isRelationship = OpcPartName.TryGetRelationshipSource(name, out var source);
            if (isRelationship != contentType.Equals(OpcMetadata.RelationshipsContentType, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException($"Invalid OPC relationship entry '{name}'.");

            if (isRelationship)
                relationshipEntries.Add((source, entry));
            else
                _parts.Add(partName, new OpcPart(this, partName, contentType, entry));
        }

        foreach (var (source, entry) in relationshipEntries)
        {
            var relationships = source switch
            {
                null => Relationships,
                _ when _parts.TryGetValue(source, out var part) => part.Relationships,
                _ => throw new InvalidDataException($"OPC relationship source '{source}' does not exist."),
            };
            OpcMetadata.ReadRelationships(entry, options.MaxMetadataCharacters, relationships);
        }
    }

    private static string GetExtension(string partName)
    {
        var dot = partName.LastIndexOf('.');
        return dot > partName.LastIndexOf('/') ? partName[(dot + 1)..] : string.Empty;
    }

    private static void RemoveLinks(OpcRelationshipCollection relationships, string partName)
    {
        foreach (var relationship in relationships.Where(relationship => string.Equals(partName, relationship.TargetPartName, StringComparison.OrdinalIgnoreCase)).ToList())
            relationships.Remove(relationship.Id);
    }
}
