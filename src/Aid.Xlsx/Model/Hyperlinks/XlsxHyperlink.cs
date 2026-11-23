#nullable disable

using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Model.Hyperlinks;

public partial class XlsxHyperlink
{
    private Uri _externalAddress;
    private string _internalAddress;

    public XlsxHyperlink(string address)
    {
        SetValues(address, string.Empty);
    }

    public XlsxHyperlink(string address, string tooltip)
    {
        SetValues(address, tooltip);
    }

    public XlsxHyperlink(IXlsxCell cell)
    {
        SetValues(cell, string.Empty);
    }

    public XlsxHyperlink(IXlsxCell cell, string tooltip)
    {
        SetValues(cell, tooltip);
    }

    public XlsxHyperlink(IXlsxRangeBase range)
    {
        SetValues(range, string.Empty);
    }

    public XlsxHyperlink(IXlsxRangeBase range, string tooltip)
    {
        SetValues(range, tooltip);
    }

    public XlsxHyperlink(Uri uri)
    {
        SetValues(uri, string.Empty);
    }

    public XlsxHyperlink(Uri uri, string tooltip)
    {
        SetValues(uri, tooltip);
    }

    public bool IsExternal { get; set; }

    public Uri ExternalAddress
    {
        get
        {
            return IsExternal ? _externalAddress : null;
        }
        set
        {
            _externalAddress = value;
            IsExternal = true;
        }
    }

#nullable enable
    /// <summary>
    /// Gets top left cell of a hyperlink range. Return <c>null</c>,
    /// if the hyperlink isn't in a worksheet.
    /// </summary>
    public IXlsxCell? Cell
    {
        get
        {
            if (Container is null)
                return null;

            return Container.GetCell(this);
        }
    }
#nullable disable

    public string InternalAddress
    {
        get
        {
            if (IsExternal)
                return null;
            if (_internalAddress.Contains('!'))
            {
                return _internalAddress[0] != '\''
                           ? string.Concat(
                                _internalAddress
[.._internalAddress.IndexOf('!')]
                                    .EscapeSheetName(),
                                '!',
                                _internalAddress[(_internalAddress.IndexOf('!') + 1)..])
                           : _internalAddress;
            }

            if (Container is null)
                throw new InvalidOperationException("Hyperlink is not attached to a worksheet.");

            var sheetName = Container.WorksheetName;
            return string.Concat(
                sheetName.EscapeSheetName(),
                '!',
                _internalAddress);
        }
        set
        {
            _internalAddress = value;
            IsExternal = false;
        }
    }

    /// <summary>
    /// Tooltip displayed when user hovers over the hyperlink range. If not specified,
    /// the link target is displayed in the tooltip.
    /// </summary>
    public string Tooltip { get; set; }

    /// <inheritdoc cref="IXlsxHyperlinks.Delete(XlsxHyperlink)"/>
    public void Delete()
    {
        Container?.Delete(this);
    }
}
