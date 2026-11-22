namespace Aid.Xlsx.Styles.Formatting;

internal record XlsxProtectionFormatValue
{
    /// <summary>
    /// Default values of protection properties in XML part. If a value is missing in XML this one is used instead.
    /// </summary>
    internal static XlsxProtectionFormatValue Default { get; } = new()
    {
        Locked = true,
        Hidden = false
    };

    public required bool Locked { get; init; }

    public required bool Hidden { get; init; }
}
