using System.Globalization;
using NUnit.Framework;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Tests.Excel.CalcEngine;

[TestFixture]
public class CalcEngineExceptionTests
{
    [OneTimeSetUp]
    public void SetCultureInfo()
    {
        Thread.CurrentThread.CurrentCulture = CultureInfo.CreateSpecificCulture("en-US");
    }

    [Test]
    public void InvalidCharNumber()
    {
        Assert.AreEqual(XlsxError.IncompatibleValue, XlsxWorkbook.EvaluateExpr("CHAR(-2)"));
        Assert.AreEqual(XlsxError.IncompatibleValue, XlsxWorkbook.EvaluateExpr("CHAR(270)"));
    }

    [Test]
    public void DivisionByZero()
    {
        Assert.AreEqual(XlsxError.DivisionByZero, XlsxWorkbook.EvaluateExpr("0/0"));
        Assert.AreEqual(XlsxError.DivisionByZero, new XlsxWorkbook().AddWorksheet().Evaluate("0/0"));
    }

    [Test]
    public void InvalidFunction()
    {
        Assert.AreEqual(XlsxError.NameNotRecognized, XlsxWorkbook.EvaluateExpr("XXX(A1:A2)"));

        var ws = new XlsxWorkbook().AddWorksheet();
        Assert.AreEqual(XlsxError.NameNotRecognized, ws.Evaluate("XXX(A1:A2)"));
    }

    [Test]
    public void NestedNameNotRecognizedException()
    {
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();
        ws.Cell("A1").SetFormulaA1("=XXX");
        ws.Cell("A2").SetFormulaA1(@"=IFERROR(A1, ""Success"")");

        Assert.AreEqual("Success", ws.Cell("A2").Value);
    }
}
