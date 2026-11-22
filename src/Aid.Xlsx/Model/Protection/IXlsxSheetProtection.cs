#nullable disable

// Keep this file CodeMaid organised and cleaned
using static Aid.Xlsx.Protection.XlsxProtectionAlgorithm;

namespace Aid.Xlsx.Protection
{
    public interface IXlsxSheetProtection : IXlsxElementProtection<XlsxSheetProtectionElements>
    {
        IXlsxSheetProtection Protect(XlsxSheetProtectionElements allowedElements);

        IXlsxSheetProtection Protect(Algorithm algorithm, XlsxSheetProtectionElements allowedElements);

        IXlsxSheetProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm, XlsxSheetProtectionElements allowedElements = XlsxSheetProtectionElements.SelectEverything);
    }
}
