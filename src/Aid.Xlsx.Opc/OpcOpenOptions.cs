namespace Aid.Xlsx.Opc;

/// <summary>Resource limits for opening an OPC container.</summary>
public sealed class OpcOpenOptions
{
    public int MaxEntries { get; init; } = 100_000;

    public long MaxUncompressedBytes { get; init; } = 4L * 1024 * 1024 * 1024;

    public double MaxCompressionRatio { get; init; } = 1000;

    public long MaxMetadataCharacters { get; init; } = 16 * 1024 * 1024;
}
