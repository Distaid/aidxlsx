#nullable disable

namespace Aid.Xlsx.Model.Drawings;

public interface IXlsxPictures : IEnumerable<IXlsxPicture>
{
    int Count { get; }

    IXlsxPicture Add(Stream stream);

    IXlsxPicture Add(Stream stream, string name);

    IXlsxPicture Add(Stream stream, XlsxPictureFormat format);

    IXlsxPicture Add(Stream stream, XlsxPictureFormat format, string name);

    IXlsxPicture Add(string imageFile);

    IXlsxPicture Add(string imageFile, string name);

    bool Contains(string pictureName);

    void Delete(string pictureName);

    void Delete(IXlsxPicture picture);

    IXlsxPicture Picture(string pictureName);

    bool TryGetPicture(string pictureName, out IXlsxPicture picture);
}
