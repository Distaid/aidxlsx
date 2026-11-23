#nullable disable

using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;
using Aid.Xlsx.Model.Workbooks;

namespace Aid.Xlsx.Model.Hyperlinks;

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

    internal void SetValues(string address, string tooltip)
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

    internal void SetValues(Uri uri, string tooltip)
    {
        Tooltip = tooltip;
        _externalAddress = uri;
        IsExternal = true;
    }

    internal void SetValues(IXlsxCell cell, string tooltip)
    {
        Tooltip = tooltip;
        _internalAddress = cell.Address.ToString(XlsxReferenceStyle.A1, true);
        IsExternal = false;
    }

    internal void SetValues(IXlsxRangeBase range, string tooltip)
    {
        Tooltip = tooltip;
        _internalAddress = range.RangeAddress.ToString(XlsxReferenceStyle.A1, true);
        IsExternal = false;
    }

    internal XlsxHyperlinks Container { get; set; }
}
