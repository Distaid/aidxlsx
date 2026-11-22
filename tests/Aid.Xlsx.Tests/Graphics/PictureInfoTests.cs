using System.Drawing;
using System.Reflection;
using NUnit.Framework;
using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Rendering;
using Aid.Xlsx.Rendering.Images;

namespace Aid.Xlsx.Tests.Graphics;

[TestFixture]
public class PictureInfoTests
{
    [Test]
    public void CanReadPng()
    {
        AssertRasterImage("SampleImagePng.png", XlsxPictureFormat.Png, new Size(252, 152), 96, 96);
    }

    [TestCase("SampleImageJfif.jpg", 176, 270, 96, 96)]
    [TestCase("jpeg-rgb.jpg", 200, 200, 0, 0)] // Adobe JPG, has APP14 marker right after SOI instead of APP0
    public void CanReadJfif(string filename, int widthPx, int heightPx, int dpiX, int dpiY)
    {
        AssertRasterImage($"Jpg.{filename}", XlsxPictureFormat.Jpeg, new Size(widthPx, heightPx), dpiX, dpiY);
    }

    [Test]
    public void CanReadExif()
    {
        AssertRasterImage("SampleImageExif.jpg", XlsxPictureFormat.Jpeg, new Size(252, 152), 0, 0);
    }

    [Test]
    public void CanReadGif87Image()
    {
        AssertRasterImage("SampleImageGif87a.gif", XlsxPictureFormat.Gif, new Size(500, 200), 0, 0);
    }

    [Test]
    public void CanReadGif89Image()
    {
        AssertRasterImage("SampleImageGif89a.gif", XlsxPictureFormat.Gif, new Size(500, 200), 0, 0);
    }

    [TestCase("SampleImageBmpWin24bit.bmp")]
    [TestCase("SampleImageBmpWin8bit.bmp")]
    [TestCase("SampleImageBmpWin4bit.bmp")]
    [TestCase("SampleImageBmpWin24bit.bmp")]
    public void CanReadBmpImageV3AndFurther(string imageName)
    {
        AssertRasterImage(imageName, XlsxPictureFormat.Bmp, new Size(167, 51), 80.645d, 80.645d);
    }

    [Test]
    public void CanReadBmpV1()
    {
        AssertRasterImage("SampleImageBmpV1.bmp", XlsxPictureFormat.Bmp, new Size(150, 50), 0, 0);
    }

    [Test]
    public void CanReadTiffWithBigEndianEncoding()
    {
        AssertRasterImage("SampleImageTiffBigEndian.tiff", XlsxPictureFormat.Tiff, new Size(130, 45), 96, 96);
    }

    [Test]
    public void CanReadTiffWithLittleEndianEncoding()
    {
        AssertRasterImage("SampleImageTiffLittleEndian.tiff", XlsxPictureFormat.Tiff, new Size(130, 45), 96, 96);
    }

    [Test]
    public void CanReadPcx()
    {
        AssertRasterImage("SampleImagePcx.pcx", XlsxPictureFormat.Pcx, new Size(100, 50), 96, 96);
    }

    [Test]
    public void CanReadWmfWithPlaceableHeader()
    {
        AssertVectorImage("SampleImagePlaceableWmf.wmf", XlsxPictureFormat.Wmf, new Size(1000, 500));
    }

    [Test]
    public void CanReadWmfWithOriginalHeader()
    {
        AssertVectorImage("SampleImageOriginalWmf.wmf", XlsxPictureFormat.Wmf, new Size(12496, 6247));
    }

    [Test]
    public void CanReadEmf()
    {
        AssertVectorImage("SampleImageEmf.emf", XlsxPictureFormat.Emf, new Size(28844, 28938));
    }

    [Test]
    public void CanReadExtendedWebp()
    {
        AssertRasterImage("SampleImageWebpExtendedFormat.webp", XlsxPictureFormat.Webp, new Size(188, 231), 72, 72);
    }

    [Test]
    public void CanReadLossyWebp()
    {
        AssertRasterImage("SampleImageWebpLossy.webp", XlsxPictureFormat.Webp, new Size(278, 90), 72, 72);
    }

    [Test]
    public void CanReadLosslessWebp()
    {
        AssertRasterImage("SampleImageWebpLossless.webp", XlsxPictureFormat.Webp, new Size(395, 136), 72, 72);
    }

    private static void AssertRasterImage(string imageName, XlsxPictureFormat expectedFormat, Size expectedPxSize, double expectedDpiX, double expectedDpiY)
    {
        AssertImage(imageName, expectedFormat, expectedPxSize, Size.Empty, expectedDpiX, expectedDpiY);
    }

    private static void AssertVectorImage(string imageName, XlsxPictureFormat expectedFormat, Size expectedHiMetricSize)
    {
        AssertImage(imageName, expectedFormat, Size.Empty, expectedHiMetricSize, 0, 0);
    }

    private static void AssertImage(string imageName, XlsxPictureFormat expectedFormat, Size expectedPxSize, Size expectedHiMetricSize, double expectedDpiX, double expectedDpiY)
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream($"Aid.Xlsx.Tests.Resource.Images.{imageName}");
        var info = DefaultGraphicEngine.Instance.Value.GetPictureInfo(stream, XlsxPictureFormat.Unknown);

        Assert.AreEqual(expectedFormat, info.Format);
        Assert.AreEqual(expectedPxSize, info.SizePx);
        Assert.AreEqual(expectedHiMetricSize, info.SizePhys);

        // Some DPI is stored as pixels per meter, causing a rounding errors.
        Assert.AreEqual(expectedDpiX, info.DpiX, 0.02);
        Assert.AreEqual(expectedDpiY, info.DpiY, 0.02);
    }
}
