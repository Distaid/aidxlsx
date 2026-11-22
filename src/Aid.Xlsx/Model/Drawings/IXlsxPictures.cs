#nullable disable

namespace Aid.Xlsx.Model.Drawings;

public interface IXlsxPictures : IEnumerable<IXlsxPicture>
{
    int Count { get; }

    IXlsxPicture Add(Stream stream);

    IXlsxPicture Add(Stream stream, String name);

    IXlsxPicture Add(Stream stream, XlsxPictureFormat format);

    IXlsxPicture Add(Stream stream, XlsxPictureFormat format, String name);

    IXlsxPicture Add(String imageFile);

    IXlsxPicture Add(String imageFile, String name);

    bool Contains(String pictureName);

    void Delete(String pictureName);

    void Delete(IXlsxPicture picture);

    IXlsxPicture Picture(String pictureName);

    bool TryGetPicture(String pictureName, out IXlsxPicture picture);
}
