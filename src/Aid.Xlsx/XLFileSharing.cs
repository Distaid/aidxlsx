// Keep this file CodeMaid organised and cleaned

namespace Aid.Xlsx
{
    internal class XLFileSharing : IXLFileSharing
    {
        public bool ReadOnlyRecommended { get; set; }
        public string? UserName { get; set; }
    }
}
