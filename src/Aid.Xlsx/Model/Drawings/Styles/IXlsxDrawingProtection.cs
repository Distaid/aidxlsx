#nullable disable

namespace Aid.Xlsx.Model.Drawings.Styles;

public interface IXlsxDrawingProtection
{
    Boolean Locked { get; set; }
    Boolean LockText { get; set; }

    IXlsxDrawingStyle SetLocked(); IXlsxDrawingStyle SetLocked(Boolean value);
    IXlsxDrawingStyle SetLockText(); IXlsxDrawingStyle SetLockText(Boolean value);

}
