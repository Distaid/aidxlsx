#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingProtection(IXlsxDrawingStyle style) : IXlsxDrawingProtection
{
    private readonly IXlsxDrawingStyle _style = style;

    public Boolean Locked { get; set; }
    public IXlsxDrawingStyle SetLocked() { Locked = true; return _style; }
    public IXlsxDrawingStyle SetLocked(Boolean value) { Locked = value; return _style; }
    public Boolean LockText { get; set; }
    public IXlsxDrawingStyle SetLockText() { LockText = true; return _style; }
    public IXlsxDrawingStyle SetLockText(Boolean value) { LockText = value; return _style; }

}
