// Keep this file CodeMaid organised and cleaned
using NUnit.Framework;
using static Aid.Xlsx.Model.Protection.XlsxProtectionAlgorithm;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Protection;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Excel.Worksheets;

[TestFixture]
public class XlsxSheetProtectionTests
{
    [Test]
    public void AllowEverything()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowedElements = XlsxSheetProtectionElements.Everything;

            foreach (var element in Enum.GetValues(typeof(XlsxSheetProtectionElements)).Cast<XlsxSheetProtectionElements>())
                Assert.IsTrue(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
        }

        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowElement(XlsxSheetProtectionElements.Everything);

            foreach (var element in Enum.GetValues(typeof(XlsxSheetProtectionElements)).Cast<XlsxSheetProtectionElements>())
                Assert.IsTrue(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
        }

        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowEverything();

            foreach (var element in Enum.GetValues(typeof(XlsxSheetProtectionElements)).Cast<XlsxSheetProtectionElements>())
                Assert.IsTrue(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
        }
    }

    [Test]
    public void AllowNothing()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowedElements = XlsxSheetProtectionElements.None;

            foreach (var element in Enum.GetValues(typeof(XlsxSheetProtectionElements))
                .Cast<XlsxSheetProtectionElements>()
                .Where(e => e != XlsxSheetProtectionElements.None))

                Assert.IsFalse(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
        }

        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            ws.Protect().AllowNone();

            foreach (var element in Enum.GetValues(typeof(XlsxSheetProtectionElements))
                .Cast<XlsxSheetProtectionElements>()
                .Where(e => e != XlsxSheetProtectionElements.None))

                Assert.IsFalse(ws.Protection.AllowedElements.HasFlag(element), element.ToString());
        }
    }

    [Test]
    public void ChangeHashingAlgorithm()
    {
        using (var ms = new MemoryStream())
        {
            using (var wb = new XlsxWorkbook())
            {
                var ws = wb.AddWorksheet();
                ws.Protect("123", Algorithm.SimpleHash);

                wb.SaveAs(ms);
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XlsxWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                Assert.IsTrue(ws.Protection.IsProtected);
                Assert.AreEqual(Algorithm.SimpleHash, ws.Protection.Algorithm);

                ws.Unprotect("123");
                ws.Protect("123", Algorithm.SHA512);
                wb.Save();
            }

            ms.Seek(0, SeekOrigin.Begin);

            using (var wb = new XlsxWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                Assert.IsTrue(ws.Protection.IsProtected);
                Assert.AreEqual(Algorithm.SHA512, ws.Protection.Algorithm);

                Assert.DoesNotThrow(() => ws.Unprotect("123"));
            }
        }
    }

    [Test]
    public void CopyProtectionFromAnotherSheet()
    {
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\Misc\SheetProtection.xlsx")))
        using (var wb = new XlsxWorkbook(stream))
        {
            var ws1 = wb.Worksheet("Protected Password = 123");
            var p1 = ws1.Protection.CastTo<XlsxSheetProtection>();
            Assert.IsTrue(p1.IsProtected);

            var ws2 = ws1.CopyTo("New worksheet");
            Assert.IsFalse(ws2.Protection.IsProtected);
            var p2 = ws2.Protection.CopyFrom(p1).CastTo<XlsxSheetProtection>();

            Assert.IsTrue(p2.IsProtected);
            Assert.IsTrue(p2.IsPasswordProtected);
            Assert.AreEqual(p1.Algorithm, p2.Algorithm);
            Assert.AreEqual(p1.PasswordHash, p2.PasswordHash);
            Assert.AreEqual(p1.Base64EncodedSalt, p2.Base64EncodedSalt);
            Assert.AreEqual(p1.SpinCount, p2.SpinCount);

            Assert.IsTrue(p2.AllowedElements.HasFlag(XlsxSheetProtectionElements.InsertColumns));
            Assert.IsTrue(p2.AllowedElements.HasFlag(XlsxSheetProtectionElements.InsertRows));
            Assert.IsFalse(p2.AllowedElements.HasFlag(XlsxSheetProtectionElements.InsertHyperlinks));

            Assert.Throws<InvalidOperationException>(() => ws2.Unprotect());
            ws2.Unprotect("123");
        }
    }

    [Test]
    public void SetWorksheetProtectionCloning()
    {
        var ws1 = new XlsxWorkbook().AddWorksheet();
        var ws2 = new XlsxWorkbook().AddWorksheet();

        ws1.Protect("123")
            .AllowElement(XlsxSheetProtectionElements.FormatEverything)
            .DisallowElement(XlsxSheetProtectionElements.FormatCells);

        Assert.AreEqual(XlsxSheetProtectionElements.FormatColumns | XlsxSheetProtectionElements.FormatRows | XlsxSheetProtectionElements.SelectEverything, ws1.Protection.AllowedElements);

        ws2.Protection = ws1.Protection;

        Assert.IsFalse(ReferenceEquals(ws1.Protection, ws2.Protection));
        Assert.IsTrue(ws2.Protection.IsProtected);
        Assert.AreEqual(XlsxSheetProtectionElements.FormatColumns | XlsxSheetProtectionElements.FormatRows | XlsxSheetProtectionElements.SelectEverything, ws2.Protection.AllowedElements);
        Assert.AreEqual((ws1.Protection as XlsxSheetProtection).PasswordHash, (ws2.Protection as XlsxSheetProtection).PasswordHash);
    }

    [Test]
    public void TestUnprotectWorksheetWithNoPassword()
    {
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx")))
        using (var wb = new XlsxWorkbook(stream))
        {
            var ws = wb.Worksheet("Sheet1");
            Assert.IsTrue(ws.Protection.IsProtected);
            ws.Unprotect();
            Assert.IsFalse(ws.Protection.IsProtected);
        }
    }

    [Test]
    public void TestWorksheetWithSHA512Protection()
    {
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"TryToLoad\SHA512PasswordProtection.xlsx")))
        using (var wb = new XlsxWorkbook(stream))
        {
            var ws = wb.Worksheet("Sheet2");
            Assert.IsTrue(ws.Protection.IsProtected);

            // Password required
            Assert.Throws<InvalidOperationException>(() => ws.Unprotect());

            Assert.AreEqual(Algorithm.SHA512, ws.Protection.Algorithm);
            ws.Unprotect("abc");
            Assert.IsFalse(ws.Protection.IsProtected);
        }
    }
}
