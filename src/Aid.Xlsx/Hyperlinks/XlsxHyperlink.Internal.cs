#nullable disable

using Aid.Xlsx.Cells;
using Aid.Xlsx.Ranges;
using Aid.Xlsx.Workbooks;

namespace Aid.Xlsx.Hyperlinks
{
    public partial class XlsxHyperlink
    {
        internal XlsxHyperlink()
        {

        }

        internal XlsxHyperlink(XlsxHyperlink hyperlink)
        {
            _externalAddress = hyperlink._externalAddress;
            _internalAddress = hyperlink._internalAddress;
            Tooltip = hyperlink.Tooltip;
            IsExternal = hyperlink.IsExternal;
        }

        internal void SetValues(String address, String tooltip)
        {
            Tooltip = tooltip;
            if (address[0] == '.')
            {
                _externalAddress = new Uri(address, UriKind.Relative);
                IsExternal = true;
            }
            else
            {
                if (Uri.TryCreate(address, UriKind.Absolute, out Uri uri))
                {
                    _externalAddress = uri;
                    IsExternal = true;
                }
                else
                {
                    _internalAddress = address;
                    IsExternal = false;
                }
            }
        }

        internal void SetValues(Uri uri, String tooltip)
        {
            Tooltip = tooltip;
            _externalAddress = uri;
            IsExternal = true;
        }

        internal void SetValues(IXlsxCell cell, String tooltip)
        {
            Tooltip = tooltip;
            _internalAddress = cell.Address.ToString(XlsxReferenceStyle.A1, true);
            IsExternal = false;
        }

        internal void SetValues(IXlsxRangeBase range, String tooltip)
        {
            Tooltip = tooltip;
            _internalAddress = range.RangeAddress.ToString(XlsxReferenceStyle.A1, true);
            IsExternal = false;
        }

        internal XlsxHyperlinks Container { get; set; }
    }
}
