// Keep this file CodeMaid organised and cleaned

namespace Aid.Xlsx.Model.Workbooks;

internal class XlsxFileSharing : IXlsxFileSharing
{
    public bool ReadOnlyRecommended { get; set; }
    public string? UserName { get; set; }
}
