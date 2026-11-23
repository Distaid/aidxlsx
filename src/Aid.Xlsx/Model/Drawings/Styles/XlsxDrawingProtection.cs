#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

internal class XlsxDrawingProtection(IXlsxDrawingStyle style) : IXlsxDrawingProtection
{
    private readonly IXlsxDrawingStyle _style = style;

    public bool Locked { get; set; }
    public IXlsxDrawingStyle SetLocked() { Locked = true; return _style; }
    public IXlsxDrawingStyle SetLocked(bool value) { Locked = value; return _style; }
    public bool LockText { get; set; }
    public IXlsxDrawingStyle SetLockText() { LockText = true; return _style; }
    public IXlsxDrawingStyle SetLockText(bool value) { LockText = value; return _style; }

}
