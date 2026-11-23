#nullable disable

using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;
using Aid.Xlsx.Model.Grid.Ranges;

namespace Aid.Xlsx.Model.Hyperlinks;

public partial class XlsxHyperlink
{
    private Uri _externalAddress;
    private String _internalAddress;

    public XlsxHyperlink(String address)
    {
        SetValues(address, String.Empty);
    }

    public XlsxHyperlink(String address, String tooltip)
    {
        SetValues(address, tooltip);
    }

    public XlsxHyperlink(IXlsxCell cell)
    {
        SetValues(cell, String.Empty);
    }

    public XlsxHyperlink(IXlsxCell cell, String tooltip)
    {
        SetValues(cell, tooltip);
    }

    public XlsxHyperlink(IXlsxRangeBase range)
    {
        SetValues(range, String.Empty);
    }

    public XlsxHyperlink(IXlsxRangeBase range, String tooltip)
    {
        SetValues(range, tooltip);
    }

    public XlsxHyperlink(Uri uri)
    {
        SetValues(uri, String.Empty);
    }

    public XlsxHyperlink(Uri uri, String tooltip)
    {
        SetValues(uri, tooltip);
    }

    public Boolean IsExternal { get; set; }

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

    public String InternalAddress
    {
        get
        {
            if (IsExternal)
                return null;
            if (_internalAddress.Contains('!'))
            {
                return _internalAddress[0] != '\''
                           ? String.Concat(
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
            return String.Concat(
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
    public String Tooltip { get; set; }

    /// <inheritdoc cref="IXlsxHyperlinks.Delete(XlsxHyperlink)"/>
    public void Delete()
    {
        Container?.Delete(this);
    }
}
