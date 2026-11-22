using System.Reflection;
using System.Xml.Linq;
using NUnit.Framework;
using SaveOptions = Aid.Xlsx.Model.Workbooks.SaveOptions;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.Workbooks;
using Aid.Xlsx.Model.Worksheets;
using Aid.Xlsx.Opc.Packaging;

namespace Aid.Xlsx.Tests.Model.Drawings;

[TestFixture]
public class PictureTests
{
    [Test]
    public void CanAddPictureFromStream()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            using (var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.SampleImageExif.jpg"))
            {
                var picture = ws.AddPicture(resourceStream, "MyPicture")
                    .WithPlacement(XlsxPicturePlacement.FreeFloating)
                    .MoveTo(50, 50)
                    .WithSize(200, 200);

                Assert.AreEqual(XlsxPictureFormat.Jpeg, picture.Format);
                Assert.AreEqual(200, picture.Width);
                Assert.AreEqual(200, picture.Height);
            }
        }
    }

    [Test]
    public void CanAddPictureFromFile()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

            try
            {
                using (var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.SampleImageExif.jpg"))
                using (var fileStream = File.Create(path))
                {
                    resourceStream.Seek(0, SeekOrigin.Begin);
                    resourceStream.CopyTo(fileStream);
                    fileStream.Close();
                }

                var picture = ws.AddPicture(path)
                    .WithPlacement(XlsxPicturePlacement.FreeFloating)
                    .MoveTo(50, 50);

                Assert.AreEqual(XlsxPictureFormat.Jpeg, picture.Format);
                Assert.AreEqual(252, picture.Width);
                Assert.AreEqual(152, picture.Height);
            }
            finally
            {
                if (File.Exists(path))
                    File.Delete(path);
            }
        }
    }

    [Test]
    public void CanAddPictureConcurrentlyFromFile()
    {
        var path = Path.ChangeExtension(Path.GetTempFileName(), "jpg");

        try
        {
            using (var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.SampleImageExif.jpg"))
            using (var fileStream = File.Create(path))
            {
                resourceStream.Seek(0, SeekOrigin.Begin);
                resourceStream.CopyTo(fileStream);
                fileStream.Close();
            }

            Parallel.Invoke(() => verifyAddImageFromFile(path), () => verifyAddImageFromFile(path));
        }
        finally
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
    }

    private void verifyAddImageFromFile(string filePath)
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            var picture = ws.AddPicture(filePath)
                       .WithPlacement(XlsxPicturePlacement.FreeFloating)
                       .MoveTo(50, 50);

            Assert.AreEqual(XlsxPictureFormat.Jpeg, picture.Format);
            Assert.AreEqual(252, picture.Width);
            Assert.AreEqual(50, picture.Top);
        }
    }

    [Test]
    public void CanScaleImage()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            using (var resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png"))
            {
                var pic = ws.AddPicture(resourceStream, "MyPicture")
                    .WithPlacement(XlsxPicturePlacement.FreeFloating)
                    .MoveTo(50, 50);

                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(252, pic.Width);
                Assert.AreEqual(152, pic.Height);

                pic.ScaleHeight(0.7);
                pic.ScaleWidth(1.2);

                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(302, pic.Width);
                Assert.AreEqual(106, pic.Height);

                pic.ScaleHeight(0.7);
                pic.ScaleWidth(1.2);

                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(362, pic.Width);
                Assert.AreEqual(74, pic.Height);

                pic.ScaleHeight(0.8, true);
                pic.ScaleWidth(1.1, true);

                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(277, pic.Width);
                Assert.AreEqual(122, pic.Height);
            }
        }
    }

    [Test]
    public void TestDefaultPictureNames()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png"))
            {
                ws.AddPicture(stream, XlsxPictureFormat.Png);
                stream.Position = 0;

                ws.AddPicture(stream, XlsxPictureFormat.Png);
                stream.Position = 0;

                ws.AddPicture(stream, XlsxPictureFormat.Png).Name = "Picture 4";
                stream.Position = 0;

                ws.AddPicture(stream, XlsxPictureFormat.Png);
                stream.Position = 0;
            }

            Assert.AreEqual("Picture 1", ws.Pictures.Skip(0).First().Name);
            Assert.AreEqual("Picture 2", ws.Pictures.Skip(1).First().Name);
            Assert.AreEqual("Picture 4", ws.Pictures.Skip(2).First().Name);
            Assert.AreEqual("Picture 5", ws.Pictures.Skip(3).First().Name);
        }
    }

    [Test]
    public void TestDefaultIds()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");

            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png"))
            {
                ws.AddPicture(stream, XlsxPictureFormat.Png);
                stream.Position = 0;

                ws.AddPicture(stream, XlsxPictureFormat.Png);
                stream.Position = 0;

                ws.AddPicture(stream, XlsxPictureFormat.Png).Name = "Picture 4";
                stream.Position = 0;

                ws.AddPicture(stream, XlsxPictureFormat.Png);
                stream.Position = 0;
            }

            Assert.AreEqual(1, ws.Pictures.Skip(0).First().Id);
            Assert.AreEqual(2, ws.Pictures.Skip(1).First().Id);
            Assert.AreEqual(3, ws.Pictures.Skip(2).First().Id);
            Assert.AreEqual(4, ws.Pictures.Skip(3).First().Id);
        }
    }

    [Test]
    public void XlsxMarkerTests()
    {
        IXlsxWorksheet ws = new XlsxWorkbook().Worksheets.Add("Sheet1");
        XlsxMarker firstMarker = new XlsxMarker(ws.Cell(1, 10), new System.Drawing.Point(100, 0));

        Assert.AreEqual(10, firstMarker.ColumnNumber);
        Assert.AreEqual(1, firstMarker.RowNumber);
        Assert.AreEqual(100, firstMarker.Offset.X);
        Assert.AreEqual(0, firstMarker.Offset.Y);
    }

    [Test]
    public void XlsxPictureTests()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.Worksheets.Add("Sheet1");

            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png"))
            {
                var pic = ws.AddPicture(stream, XlsxPictureFormat.Png, "Image1")
                    .WithPlacement(XlsxPicturePlacement.FreeFloating)
                    .MoveTo(220, 155);

                Assert.AreEqual(XlsxPicturePlacement.FreeFloating, pic.Placement);
                Assert.AreEqual("Image1", pic.Name);
                Assert.AreEqual(XlsxPictureFormat.Png, pic.Format);
                Assert.AreEqual(252, pic.OriginalWidth);
                Assert.AreEqual(152, pic.OriginalHeight);
                Assert.AreEqual(252, pic.Width);
                Assert.AreEqual(152, pic.Height);
                Assert.AreEqual(220, pic.Left);
                Assert.AreEqual(155, pic.Top);
            }
        }
    }

    [Test]
    public void CanLoadFileWithImagesAndCopyImagesToNewSheet()
    {
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
        using (var wb = new XlsxWorkbook(stream))
        {
            var ws = wb.Worksheets.First();
            Assert.AreEqual(2, ws.Pictures.Count);

            var copy = ws.CopyTo("NewSheet");
            Assert.AreEqual(2, copy.Pictures.Count);
        }
    }

    [Test]
    public void CanDeletePictures()
    {
        using (var ms = new MemoryStream())
        {
            int originalCount;

            using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
            using (var wb = new XlsxWorkbook(stream))
            {
                var ws = wb.Worksheets.First();
                originalCount = ws.Pictures.Count;
                ws.Pictures.Delete(ws.Pictures.First());

                var pictureName = ws.Pictures.First().Name;
                ws.Pictures.Delete(pictureName);

                wb.SaveAs(ms);
            }

            using (var wb = new XlsxWorkbook(ms))
            {
                var ws = wb.Worksheets.First();
                Assert.AreEqual(originalCount - 2, ws.Pictures.Count);
            }
        }
    }

    [Test]
    public void DeletingLastImagePreservesOtherShapes()
    {
        using var input = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Pictures\ImageShapeZOrder-Input.xlsx"));
        using var workbook = new XlsxWorkbook(input);
        var sheet = workbook.Worksheets.Single();
        sheet.Pictures.Delete(sheet.Pictures.Single());
        using var output = new MemoryStream();
        workbook.SaveAs(output, new SaveOptions { ValidatePackage = true });
        using var package = OpcPackage.Open(output);
        var drawing = package.GetPart("/xl/drawings/drawing1.xml");
        using var drawingStream = drawing.OpenRead();
        var root = XElement.Load(drawingStream);
        var ns = root.Name.Namespace;
        Assert.AreEqual(2, root.Descendants(ns + "sp").Count());
        Assert.IsEmpty(root.Descendants(ns + "pic"));
        Assert.IsEmpty(drawing.Relationships);
    }

    [Test]
    public void PictureRenameTests()
    {
        using (var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Examples\ImageHandling\ImageAnchors.xlsx")))
        using (var wb = new XlsxWorkbook(stream))
        {
            var ws = wb.Worksheet("Images3");
            var picture = ws.Pictures.First();
            Assert.AreEqual("Picture 1", picture.Name);

            picture.Name = "picture 1";
            picture.Name = "pICture 1";
            picture.Name = "Picture 1";

            picture = ws.Pictures.Last();
            picture.Name = "new name";

            Assert.Throws<ArgumentException>(() => picture.Name = "Picture 1");
            Assert.Throws<ArgumentException>(() => picture.Name = "picTURE 1");
        }
    }

    [Test]
    public void HandleDuplicatePictureIdsAcrossWorksheets()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws1 = wb.AddWorksheet("Sheet1");
            var ws2 = wb.AddWorksheet("Sheet2");

            using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png"))
            {
                (ws1 as XlsxWorksheet).AddPicture(stream, "Picture 1", 2);
                (ws1 as XlsxWorksheet).AddPicture(stream, "Picture 2", 3);

                //Internal method - used for loading files
                var pic = (ws2 as XlsxWorksheet).AddPicture(stream, "Picture 1", 2)
                    .WithPlacement(XlsxPicturePlacement.FreeFloating)
                    .MoveTo(220, 155) as XlsxPicture;

                var id = pic.Id;

                pic.Id = id;
                Assert.AreEqual(id, pic.Id);

                pic.Id = 3;
                Assert.AreEqual(3, pic.Id);

                pic.Id = id;

                var pic2 = (ws2 as XlsxWorksheet).AddPicture(stream, "Picture 2", 3)
                    .WithPlacement(XlsxPicturePlacement.FreeFloating)
                    .MoveTo(440, 300) as XlsxPicture;
            }
        }
    }

    [Test]
    public void CopyImageSameWorksheet()
    {
        var wb = new XlsxWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");

        IXlsxPicture original;
        using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png"))
        {
            original = (ws1 as XlsxWorksheet).AddPicture(stream, "Picture 1", 2)
                .WithPlacement(XlsxPicturePlacement.FreeFloating)
                .MoveTo(220, 155) as XlsxPicture;
        }

        var copy = original.Duplicate()
            .MoveTo(300, 200) as XlsxPicture;

        Assert.AreEqual(2, ws1.Pictures.Count());
        Assert.AreEqual(ws1, copy.Worksheet);
        Assert.AreEqual(original.Format, copy.Format);
        Assert.AreEqual(original.Height, copy.Height);
        Assert.AreEqual(original.Placement, copy.Placement);
        Assert.AreEqual(original.TopLeftCell.ToString(), copy.TopLeftCell.ToString());
        Assert.AreEqual(original.Width, copy.Width);
        Assert.AreEqual(original.ImageStream.ToArray(), copy.ImageStream.ToArray(), "Image streams differ");

        Assert.AreEqual(200, copy.Top);
        Assert.AreEqual(300, copy.Left);
        Assert.AreNotEqual(original.Id, copy.Id);
        Assert.AreNotEqual(original.Name, copy.Name);
    }

    [Test]
    public void CopyImageDifferentWorksheets()
    {
        var wb = new XlsxWorkbook();
        var ws1 = wb.Worksheets.Add("Sheet1");
        IXlsxPicture original;
        using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png"))
        {
            original = (ws1 as XlsxWorksheet).AddPicture(stream, "Picture 1", 2)
                .WithPlacement(XlsxPicturePlacement.FreeFloating)
                .MoveTo(220, 155) as XlsxPicture;
        }
        var ws2 = wb.Worksheets.Add("Sheet2");

        var copy = original.CopyTo(ws2);

        Assert.AreEqual(1, ws1.Pictures.Count());
        Assert.AreEqual(1, ws2.Pictures.Count());

        Assert.AreEqual(ws2, copy.Worksheet);

        Assert.AreEqual(original.Format, copy.Format);
        Assert.AreEqual(original.Height, copy.Height);
        Assert.AreEqual(original.Left, copy.Left);
        Assert.AreEqual(original.Name, copy.Name);
        Assert.AreEqual(original.Placement, copy.Placement);
        Assert.AreEqual(original.Top, copy.Top);
        Assert.AreEqual(original.TopLeftCell.ToString(), copy.TopLeftCell.ToString());
        Assert.AreEqual(original.Width, copy.Width);
        Assert.AreEqual(original.ImageStream.ToArray(), copy.ImageStream.ToArray(), "Image streams differ");

        Assert.AreNotEqual(original.Id, copy.Id);
    }

    [Test]
    public void PictureShiftsWhenInsertingRows()
    {
        using (var wb = new XlsxWorkbook())
        using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png"))
        {
            var ws = wb.Worksheets.Add("ImageShift");
            var picture = ws.AddPicture(stream, XlsxPictureFormat.Png, "PngImage")
                .MoveTo(ws.Cell(5, 2))
                .WithPlacement(XlsxPicturePlacement.Move);

            ws.Row(2).InsertRowsBelow(20);

            Assert.AreEqual(25, picture.TopLeftCell.Address.RowNumber);
        }
    }

    [Test]
    public void PictureNotFound()
    {
        using (var wb = new XlsxWorkbook())
        {
            var ws = wb.AddWorksheet("Sheet1");
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Picture("dummy"));
            Assert.Throws<ArgumentOutOfRangeException>(() => ws.Pictures.Delete("dummy"));
        }
    }

    [Test]
    public void CanCopyEmfPicture()
    {
        // #1621 - There are 2 Bmp Guids: ImageFormat.Bmp and ImageFormat.MemoryBmp
        using var stream = TestHelper.GetStreamFromResource(TestHelper.GetResourcePath(@"Other\Pictures\EmfPicture.xlsx"));
        using var wb = new XlsxWorkbook(stream);
        var ws1 = wb.Worksheets.First();
        var img1 = ws1.Pictures.First();

        var ws2 = wb.AddWorksheet();

        var img2 = img1.CopyTo(ws2);

        Assert.AreEqual(XlsxPictureFormat.Emf, img2.Format);

        using var ms = new MemoryStream();
        wb.SaveAs(ms);

        ms.Seek(0, SeekOrigin.Begin);

        using var wb2 = new XlsxWorkbook(ms);
        ws2 = wb2.Worksheet("Sheet2");
        img2 = ws2.Pictures.First();
        Assert.AreEqual(XlsxPictureFormat.Emf, img2.Format);
    }

    [Test]
    public void KeepOriginalDrawingShapesZOrder()
    {
        // File contains shapes and a picture in a mixed order.
        TestHelper.LoadSaveAndCompare(
            @"Other\Pictures\ImageShapeZOrder-Input.xlsx",
            @"Other\Pictures\ImageShapeZOrder-Output.xlsx");
    }

    [TestCase("@")]
    [TestCase(":")]
    [TestCase("\\")]
    [TestCase("/")]
    [TestCase("?")]
    [TestCase("*")]
    [TestCase("[]")]
    [TestCase(" ")] // Whitespace name is allowed, but can't be empty
    [TestCase("C:\\Images\\pic.jpg")] // Path with multiple forbidden chars
    [TestCase("http://example.com/image.jpg")] // URL with multiple forbidden chars
    [TestCase("Picture@01\\QPosted@")] // A name from a problematic workbook
    public void PictureCanHaveUnusualCharactersInName(string nameWithUnusualCharacter)
    {
        // The name of a picture couldn't contain certain characters in some ancient version of Excel. Verify that
        // it is no longer the case through the whole lifecycle (add picture, change name, save, load).
        AssertPictureNameAllowed(nameWithUnusualCharacter);
    }

    [Test]
    public void PictureNameCanBeLong()
    {
        // Picture name was originally limited to 31 characters. Verify that it is no longer the case.
        var longName = new string('a', 100_000);
        AssertPictureNameAllowed(longName);
    }

    [TestCase(null)]
    [TestCase("")]
    public void PictureNameCantBeNullOrEmpty(string invalidName)
    {
        // Picture name is a required attribute, though Excel generates a name if it isn't specified instead of failing.
        using var wb = new XlsxWorkbook();
        var ws = wb.AddWorksheet();

        using var imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png");
        var picture = ws.AddPicture(imageStream);
        Assert.Throws<ArgumentException>(() => picture.Name = invalidName);
    }

    private static void AssertPictureNameAllowed(string testedName)
    {
        TestHelper.CreateSaveLoadAssert(
            wb =>
            {
                using var imageStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("Aid.Xlsx.Tests.Resource.Images.ImageHandling.png");
                var ws1 = wb.AddWorksheet("AddPicture");
                ws1.AddPicture(imageStream, testedName);

                var ws2 = wb.AddWorksheet("Setter");
                var picture = ws2.AddPicture(imageStream);
                picture.Name = testedName;
            }, wb =>
            {
                Assert.AreEqual(testedName, wb.Worksheet("AddPicture").Pictures.Single().Name);
                Assert.AreEqual(testedName, wb.Worksheet("Setter").Pictures.Single().Name);
            });
    }
}
