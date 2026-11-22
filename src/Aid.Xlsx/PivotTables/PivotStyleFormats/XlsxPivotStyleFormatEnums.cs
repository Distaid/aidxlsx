namespace Aid.Xlsx.PivotTables.PivotStyleFormats;

[Flags]
public enum XlsxPivotStyleFormatElement
{
    None = 0,
    Label = 1 << 1,
    Data = 1 << 2,

    All = Label | Data
}
