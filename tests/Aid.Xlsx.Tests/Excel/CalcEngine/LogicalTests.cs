using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Workbooks;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel.CalcEngine
{
    [TestFixture]
    public class LogicalTests
    {
        [Test]
        public void And_IsLogicalConjunction()
        {
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("AND(TRUE)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("AND(TRUE, TRUE)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("AND(TRUE, TRUE, TRUE)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("AND({TRUE, TRUE}, TRUE)"));

            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("AND(FALSE)"));
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("AND(TRUE, FALSE)"));
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("AND({TRUE, FALSE})"));
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("AND(TRUE, {TRUE, FALSE})"));
        }

        [TestCase("A1")]
        [TestCase("A1:A5")]
        [TestCase("(A1:A5,B1:B5)")]
        public void And_NoCollectionValues_Error(string range)
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual(XlsxError.IncompatibleValue, ws.Evaluate($"AND({range})"));
        }

        [Test]
        public void And_ScalarArgumentsCoercedFromBlankOrTextOrNumber()
        {
            // Blank evaluated to false
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("AND(IF(TRUE,,))"));

            // Number coerced to logical
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("AND(0)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("AND(0.1)"));

            // Text coerced to logical
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("AND(\"FALSE\")"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("AND(\"TRUE\")"));
        }

        [Test]
        public void And_UnconvertableScalarArgumentsSkipped()
        {
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("AND(TRUE,\"z\")"));
        }

        [Test]
        public void And_OnlyLogicalOrNumberElementsOfCollectionUsed()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            // 0 is a number and is converted to logical
            ws.Cell("A1").Value = 0;
            Assert.AreEqual(false, ws.Evaluate("AND(TRUE,A1)"));

            // false is a logical
            ws.Cell("A2").Value = false;
            Assert.AreEqual(false, ws.Evaluate("AND(TRUE,A2)"));

            // Text is not converted and thus skipped for evaluation
            ws.Cell("A3").Value = "FALSE";
            Assert.AreEqual(true, ws.Evaluate("AND(TRUE,A3)"));

            ws.Cell("A4").Value = "some text";
            Assert.AreEqual(true, ws.Evaluate("AND(TRUE,A4)"));
        }

        [Test]
        public void If_2_Params_true()
        {
            Object actual = XlsxWorkbook.EvaluateExpr(@"if(1 = 1, ""T"")");
            Assert.AreEqual("T", actual);
        }

        [Test]
        public void If_2_Params_false()
        {
            Object actual = XlsxWorkbook.EvaluateExpr(@"if(1 = 2, ""T"")");
            Assert.AreEqual(false, actual);
        }

        [Test]
        public void If_3_Params_true()
        {
            Object actual = XlsxWorkbook.EvaluateExpr(@"if(1 = 1, ""T"", ""F"")");
            Assert.AreEqual("T", actual);
        }

        [Test]
        public void If_3_Params_false()
        {
            Object actual = XlsxWorkbook.EvaluateExpr(@"if(1 = 2, ""T"", ""F"")");
            Assert.AreEqual("F", actual);
        }

        [Test]
        public void If_Comparing_Against_Empty_String()
        {
            Object actual;
            actual = XlsxWorkbook.EvaluateExpr(@"if(date(2016, 1, 1) = """", ""A"",""B"")");
            Assert.AreEqual("B", actual);

            actual = XlsxWorkbook.EvaluateExpr(@"if("""" = date(2016, 1, 1), ""A"",""B"")");
            Assert.AreEqual("B", actual);

            actual = XlsxWorkbook.EvaluateExpr(@"if("""" = 123, ""A"",""B"")");
            Assert.AreEqual("B", actual);

            actual = XlsxWorkbook.EvaluateExpr(@"if("""" = """", ""A"",""B"")");
            Assert.AreEqual("A", actual);
        }

        [Test]
        public void If_Case_Insensitivity()
        {
            Object actual;
            actual = XlsxWorkbook.EvaluateExpr(@"IF(""text""=""TEXT"", 1, 2)");
            Assert.AreEqual(1, actual);
        }

        [Test]
        public void If_CanReturnReference()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual(true, ws.Evaluate("ISREF(IF(TRUE, A1))"));
            Assert.AreEqual(true, ws.Evaluate("ISREF(IF(FALSE,, A1))"));
        }

        [Test]
        public void If_has_scalar_condition_and_range_values()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            ws.Cell("A1").InsertData(new[] { 1, 2, 3 });
            ws.Cell("B1").InsertData(new[] { 4, 5, 6 });
            ws.Cell("C1").InsertData(new[] { true, false, true });
            for (var row = 1; row <= 4; ++row)
                ws.Cell(row, 4).FormulaA1 = "SUM(IF(C1:C3, A1:A3, B1:B3))";

            // Condition is implicitely intersected, because it's a scalar parameter
            Assert.AreEqual(6, ws.Cell("D1").Value);
            Assert.AreEqual(15, ws.Cell("D2").Value);
            Assert.AreEqual(6, ws.Cell("D3").Value);
            Assert.AreEqual(XlsxError.IncompatibleValue, ws.Cell("D4").Value);
        }

        [Test]
        public void If_ConditionError_ReturnError()
        {
            Assert.AreEqual(XlsxError.DivisionByZero, XlsxWorkbook.EvaluateExpr(@"IF(1/0, ""T"", ""F"")"));
        }

        [Test]
        public void If_ConditionCoercedToLogical()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual("F", ws.Evaluate(@"IF(A1, ""T"", ""F"")"));

            Assert.AreEqual("T", ws.Evaluate(@"IF(""TRUE"", ""T"", ""F"")"));
            Assert.AreEqual("F", ws.Evaluate(@"IF(""FALSE"", ""T"", ""F"")"));
            Assert.AreEqual(XlsxError.IncompatibleValue, ws.Evaluate(@"IF(""text"", ""T"", ""F"")"));

            Assert.AreEqual("T", ws.Evaluate(@"IF(1, ""T"", ""F"")"));
            Assert.AreEqual("F", ws.Evaluate(@"IF(0, ""T"", ""F"")"));
        }

        [Test]
        public void If_MissingValues_ReturnBlank()
        {
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr(@"ISBLANK(IF(TRUE,,))"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr(@"ISBLANK(IF(FALSE,,))"));
        }

        [Test]
        public void IfError_FirstArgumentNonError_ReturnFirstArgument()
        {
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("ISBLANK(IFERROR(IF(TRUE,), 5))"));

            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("IFERROR(FALSE, 5)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("IFERROR(TRUE, 5)"));

            Assert.AreEqual(0.0, XlsxWorkbook.EvaluateExpr("IFERROR(0, 5)"));
            Assert.AreEqual(-2.0, XlsxWorkbook.EvaluateExpr("IFERROR(-2, 5)"));

            Assert.AreEqual(string.Empty, XlsxWorkbook.EvaluateExpr("IFERROR(\"\", 5)"));
            Assert.AreEqual("text", XlsxWorkbook.EvaluateExpr("IFERROR(\"text\", 5)"));
        }

        [Test]
        public void IfError_FirstArgumentError_ReturnSecondArgument()
        {
            Assert.AreEqual("text", XlsxWorkbook.EvaluateExpr("IFERROR(1/0, \"text\")"));

            Assert.AreEqual(XlsxError.NameNotRecognized, XlsxWorkbook.EvaluateExpr("IFERROR(#REF!, #NAME?)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("IFERROR(#NULL!, TRUE)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("ISBLANK(IFERROR(#VALUE!,IF(TRUE,)))"));
        }

        [Test]
        public void IfError_ReferenceNeverReturned()
        {
            // Unlike IF, IFERROR doesn't return reference
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual(false, ws.Evaluate("ISREF(IFERROR(#VALUE!, A1))"));
        }

        [TestCase("TRUE", false)]
        [TestCase("FALSE", true)]
        [TestCase("IF(TRUE,,)", true)] // Blank
        [TestCase("0", true)]
        [TestCase("0.1", false)]
        [TestCase("\"true\"", false)]
        [TestCase("\"false\"", true)]
        [TestCase("1/0", XlsxError.DivisionByZero)]
        public void Not(string valueFormula, object expectedResult)
        {
            Assert.AreEqual(expectedResult, XlsxWorkbook.EvaluateExpr($"NOT({valueFormula})"));
        }

        [Test]
        public void Or_IsLogicalDisjunction()
        {
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("OR(TRUE)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("OR(TRUE, TRUE)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("OR(TRUE, FALSE, TRUE)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("OR({FALSE, TRUE}, FALSE)"));

            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("OR(FALSE)"));
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("OR(FALSE, FALSE)"));
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("OR({FALSE, FALSE})"));
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("OR(FALSE, {FALSE, FALSE})"));
        }

        [TestCase("A1")]
        [TestCase("A1:A5")]
        [TestCase("(A1:A5,B1:B5)")]
        public void Or_NoCollectionValues_Error(string range)
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            Assert.AreEqual(XlsxError.IncompatibleValue, ws.Evaluate($"OR({range})"));
        }

        [Test]
        public void Or_ScalarArgumentsCoercedFromBlankOrTextOrNumber()
        {
            // Blank evaluated to false
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("OR(IF(TRUE,,))"));

            // Number coerced to logical
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("OR(0)"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("OR(0.1)"));

            // Text coerced to logical
            Assert.AreEqual(false, XlsxWorkbook.EvaluateExpr("OR(\"FALSE\")"));
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("OR(\"TRUE\")"));
        }

        [Test]
        public void Or_UnconvertableScalarArgumentsSkipped()
        {
            Assert.AreEqual(true, XlsxWorkbook.EvaluateExpr("OR(TRUE,\"z\")"));
        }

        [Test]
        public void Or_OnlyLogicalOrNumberElementsOfCollectionUsed()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();

            // 1 is a number and is converted to logical
            ws.Cell("A1").Value = 1;
            Assert.AreEqual(true, ws.Evaluate("OR(FALSE,A1)"));

            // false is a logical
            ws.Cell("A2").Value = true;
            Assert.AreEqual(true, ws.Evaluate("OR(FALSE,A2)"));

            // Text is not converted and thus skipped for evaluation
            ws.Cell("A3").Value = "TRUE";
            Assert.AreEqual(false, ws.Evaluate("OR(FALSE,A3)"));

            ws.Cell("A4").Value = "some text";
            Assert.AreEqual(false, ws.Evaluate("OR(FALSE,A4)"));
        }
    }
}
