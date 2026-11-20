#nullable disable

using System;

namespace Aid.Xlsx
{
    public interface IXLDrawingProtection
    {
        Boolean Locked { get; set; }
        Boolean LockText { get; set; }

        IXLDrawingStyle SetLocked(); IXLDrawingStyle SetLocked(Boolean value);
        IXLDrawingStyle SetLockText(); IXLDrawingStyle SetLockText(Boolean value);

    }
}
