using Aid.Xlsx.Coordinates;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;
using NUnit.Framework;

namespace Aid.Xlsx.Tests
{
    [TestFixture]
    public class XlsxAddressTests
    {
        [Test]
        public void ToStringTest()
        {
            var ws = new XlsxWorkbook().Worksheets.Add("Sheet1");
            IXlsxAddress address = ws.Cell(1, 1).Address;

            Assert.AreEqual("A1", address.ToString());
            Assert.AreEqual("A1", address.ToString(XlsxReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToString(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("A1", address.ToString(XlsxReferenceStyle.Default));
            Assert.AreEqual("Sheet1!A1", address.ToString(XlsxReferenceStyle.Default, true));

            Assert.AreEqual("A1", address.ToStringRelative());
            Assert.AreEqual("Sheet1!A1", address.ToStringRelative(true));

            Assert.AreEqual("$A$1", address.ToStringFixed());
            Assert.AreEqual("$A$1", address.ToStringFixed(XlsxReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToStringFixed(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("$A$1", address.ToStringFixed(XlsxReferenceStyle.Default));
            Assert.AreEqual("Sheet1!$A$1", address.ToStringFixed(XlsxReferenceStyle.A1, true));
            Assert.AreEqual("Sheet1!R1C1", address.ToStringFixed(XlsxReferenceStyle.R1C1, true));
            Assert.AreEqual("Sheet1!$A$1", address.ToStringFixed(XlsxReferenceStyle.Default, true));
        }

        [Test]
        public void ToStringTestWithSpace()
        {
            IXlsxWorksheet ws = new XlsxWorkbook().Worksheets.Add("Sheet 1");
            IXlsxAddress address = ws.Cell(1, 1).Address;

            Assert.AreEqual("A1", address.ToString());
            Assert.AreEqual("A1", address.ToString(XlsxReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToString(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("A1", address.ToString(XlsxReferenceStyle.Default));
            Assert.AreEqual("'Sheet 1'!A1", address.ToString(XlsxReferenceStyle.Default, true));

            Assert.AreEqual("A1", address.ToStringRelative());
            Assert.AreEqual("'Sheet 1'!A1", address.ToStringRelative(true));

            Assert.AreEqual("$A$1", address.ToStringFixed());
            Assert.AreEqual("$A$1", address.ToStringFixed(XlsxReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToStringFixed(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("$A$1", address.ToStringFixed(XlsxReferenceStyle.Default));
            Assert.AreEqual("'Sheet 1'!$A$1", address.ToStringFixed(XlsxReferenceStyle.A1, true));
            Assert.AreEqual("'Sheet 1'!R1C1", address.ToStringFixed(XlsxReferenceStyle.R1C1, true));
            Assert.AreEqual("'Sheet 1'!$A$1", address.ToStringFixed(XlsxReferenceStyle.Default, true));
        }

        [Test]
        public void InvalidAddressToStringTest()
        {
            var address = ProduceInvalidAddress();

            Assert.AreEqual("#REF!", address.ToString());
            Assert.AreEqual("#REF!", address.ToString(XlsxReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToString(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("#REF!", address.ToString(XlsxReferenceStyle.Default));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToString(XlsxReferenceStyle.Default, true));
        }

        [Test]
        public void InvalidAddressToStringFixedTest()
        {
            var address = ProduceInvalidAddress();

            Assert.AreEqual("#REF!", address.ToStringFixed());
            Assert.AreEqual("#REF!", address.ToStringFixed(XlsxReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToStringFixed(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("#REF!", address.ToStringFixed(XlsxReferenceStyle.Default));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XlsxReferenceStyle.A1, true));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XlsxReferenceStyle.R1C1, true));
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringFixed(XlsxReferenceStyle.Default, true));
        }

        [Test]
        public void InvalidAddressToStringRelativeTest()
        {
            var address = ProduceInvalidAddress();

            Assert.AreEqual("#REF!", address.ToStringRelative());
            Assert.AreEqual("'Sheet 1'!#REF!", address.ToStringRelative(true));
        }

        [Test]
        public void AddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.AreEqual("A1", address.ToString());
            Assert.AreEqual("A1", address.ToString(XlsxReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToString(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("A1", address.ToString(XlsxReferenceStyle.Default));
            Assert.AreEqual("#REF!A1", address.ToString(XlsxReferenceStyle.Default, true));
        }

        [Test]
        public void AddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.AreEqual("$A$1", address.ToStringFixed());
            Assert.AreEqual("$A$1", address.ToStringFixed(XlsxReferenceStyle.A1));
            Assert.AreEqual("R1C1", address.ToStringFixed(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("$A$1", address.ToStringFixed(XlsxReferenceStyle.Default));
            Assert.AreEqual("#REF!$A$1", address.ToStringFixed(XlsxReferenceStyle.A1, true));
            Assert.AreEqual("#REF!R1C1", address.ToStringFixed(XlsxReferenceStyle.R1C1, true));
            Assert.AreEqual("#REF!$A$1", address.ToStringFixed(XlsxReferenceStyle.Default, true));
        }

        [Test]
        public void AddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceAddressOnDeletedWorksheet();

            Assert.AreEqual("A1", address.ToStringRelative());
            Assert.AreEqual("#REF!A1", address.ToStringRelative(true));
        }

        [Test]
        public void InvalidAddressOnDeletedWorksheetToStringTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!", address.ToString());
            Assert.AreEqual("#REF!", address.ToString(XlsxReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToString(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("#REF!", address.ToString(XlsxReferenceStyle.Default));
            Assert.AreEqual("#REF!#REF!", address.ToString(XlsxReferenceStyle.Default, true));
        }

        [Test]
        public void InvalidAddressOnDeletedWorksheetToStringFixedTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!", address.ToStringFixed());
            Assert.AreEqual("#REF!", address.ToStringFixed(XlsxReferenceStyle.A1));
            Assert.AreEqual("#REF!", address.ToStringFixed(XlsxReferenceStyle.R1C1));
            Assert.AreEqual("#REF!", address.ToStringFixed(XlsxReferenceStyle.Default));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XlsxReferenceStyle.A1, true));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XlsxReferenceStyle.R1C1, true));
            Assert.AreEqual("#REF!#REF!", address.ToStringFixed(XlsxReferenceStyle.Default, true));
        }

        [Test]
        public void InvalidAddressOnDeletedWorksheetToStringRelativeTest()
        {
            var address = ProduceInvalidAddressOnDeletedWorksheet();

            Assert.AreEqual("#REF!", address.ToStringRelative());
            Assert.AreEqual("#REF!#REF!", address.ToStringRelative(true));
        }

        #region Private Methods

        private IXlsxAddress ProduceInvalidAddress()
        {
            IXlsxWorksheet ws = new XlsxWorkbook().Worksheets.Add("Sheet 1");
            var range = ws.Range("A1:B2");

            ws.Rows(1, 5).Delete();
            return range.RangeAddress.FirstAddress;
        }

        private IXlsxAddress ProduceAddressOnDeletedWorksheet()
        {
            IXlsxWorksheet ws = new XlsxWorkbook().Worksheets.Add("Sheet 1");
            var address = ws.Cell("A1").Address;

            ws.Delete();
            return address;
        }

        private IXlsxAddress ProduceInvalidAddressOnDeletedWorksheet()
        {
            var address = ProduceInvalidAddress();
            address.Worksheet.Delete();
            return address;
        }

        #endregion Private Methods
    }
}
