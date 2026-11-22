using Aid.Xlsx.Workbooks;
using NUnit.Framework;

namespace Aid.Xlsx.Tests;

public class ApiSurfaceTests
{
    [Test]
    public void LibraryExportsOnlyDomainTypesWithCurrentNames()
    {
        var legacyTypes = typeof(XlsxWorkbook).Assembly.GetExportedTypes()
            .Where(type => type.Namespace == "Aid.Xlsx" ||
                           type.Name.StartsWith("XL", StringComparison.Ordinal) ||
                           type.Name.StartsWith("IXL", StringComparison.Ordinal) ||
                           type.Name.StartsWith("ClosedXML", StringComparison.Ordinal))
            .Select(type => type.FullName);

        Assert.That(legacyTypes, Is.Empty);
    }
}
