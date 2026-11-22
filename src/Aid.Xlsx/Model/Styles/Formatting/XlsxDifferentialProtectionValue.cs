namespace Aid.Xlsx.Model.Styles.Formatting;

internal record XlsxDifferentialProtectionValue
{
    internal static readonly XlsxDifferentialProtectionValue Empty = new()
    {
        Hidden = null,
        Locked = null,
    };

    internal required bool? Hidden { get; init; }

    internal required bool? Locked { get; init; }
}
