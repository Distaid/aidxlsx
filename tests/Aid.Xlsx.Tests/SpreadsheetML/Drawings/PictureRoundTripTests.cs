using NUnit.Framework;

namespace Aid.Xlsx.Tests.SpreadsheetML.Drawings;

[TestFixture]
public class PictureRoundTripTests
{
    [TestCase("Other.Drawings.picture-webp.xlsx")]
    public void Can_load_and_save_workbook_with_image_type(string resourceWithImageType)
    {
        TestHelper.LoadSaveAndCompare(resourceWithImageType, resourceWithImageType);
    }
}
