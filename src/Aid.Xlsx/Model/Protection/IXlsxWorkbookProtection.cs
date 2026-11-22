#nullable disable

// Keep this file CodeMaid organised and cleaned
using static Aid.Xlsx.Model.Protection.XlsxProtectionAlgorithm;

namespace Aid.Xlsx.Model.Protection;

public interface IXlsxWorkbookProtection : IXlsxElementProtection<XlsxWorkbookProtectionElements>
{
    IXlsxWorkbookProtection Protect(XlsxWorkbookProtectionElements allowedElements);

    IXlsxWorkbookProtection Protect(Algorithm algorithm, XlsxWorkbookProtectionElements allowedElements);

    IXlsxWorkbookProtection Protect(String password, Algorithm algorithm = DefaultProtectionAlgorithm, XlsxWorkbookProtectionElements allowedElements = XlsxWorkbookProtectionElements.Windows);
}
