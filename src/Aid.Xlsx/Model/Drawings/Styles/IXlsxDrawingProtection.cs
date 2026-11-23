#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

public interface IXlsxDrawingProtection
{
    bool Locked { get; set; }
    bool LockText { get; set; }

    IXlsxDrawingStyle SetLocked(); IXlsxDrawingStyle SetLocked(bool value);
    IXlsxDrawingStyle SetLockText(); IXlsxDrawingStyle SetLockText(bool value);

}
