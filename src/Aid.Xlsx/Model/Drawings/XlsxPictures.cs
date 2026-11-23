using System.Collections;
using System.Diagnostics;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.Drawings;

internal class XlsxPictures : IXlsxPictures, IEnumerable<XlsxPicture>
{
    private readonly List<XlsxPicture> _pictures = [];
    private readonly XlsxWorksheet _worksheet;

    public XlsxPictures(XlsxWorksheet worksheet)
    {
        _worksheet = worksheet;
        Deleted = new HashSet<String>();
    }

    public int Count
    {
        [DebuggerStepThrough]
        get { return _pictures.Count; }
    }

    internal ICollection<String> Deleted { get; private set; }

    public IXlsxPicture Add(Stream stream)
    {
        var picture = new XlsxPicture(_worksheet, stream);
        _pictures.Add(picture);
        picture.Name = GetNextPictureName();
        return picture;
    }

    public IXlsxPicture Add(Stream stream, string name)
    {
        var picture = Add(stream);
        picture.Name = name;
        return picture;
    }

    public IXlsxPicture Add(Stream stream, XlsxPictureFormat format)
    {
        var picture = new XlsxPicture(_worksheet, stream, format);
        _pictures.Add(picture);
        picture.Name = GetNextPictureName();
        return picture;
    }

    public IXlsxPicture Add(Stream stream, XlsxPictureFormat format, string name)
    {
        var picture = Add(stream, format);
        picture.Name = name;
        return picture;
    }

    public IXlsxPicture Add(string imageFile)
    {
        using var fs = File.OpenRead(imageFile);
        var picture = new XlsxPicture(_worksheet, fs);
        _pictures.Add(picture);
        picture.Name = GetNextPictureName();
        return picture;
    }

    public IXlsxPicture Add(string imageFile, string name)
    {
        var picture = Add(imageFile);
        picture.Name = name;
        return picture;
    }

    public bool Contains(string pictureName)
    {
        return _pictures.Any(p => string.Equals(p.Name, pictureName, StringComparison.OrdinalIgnoreCase));
    }

    public void Delete(IXlsxPicture picture)
    {
        Delete(picture.Name);
    }

    public void Delete(string pictureName)
    {
        var picturesToDelete = _pictures
            .Where(picture => picture.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase))
            .ToList();

        if (!picturesToDelete.Any())
            throw new ArgumentOutOfRangeException(nameof(pictureName), $"Picture {pictureName} was not found.");

        foreach (var picture in picturesToDelete)
        {
            if (!string.IsNullOrEmpty(picture.RelId))
                Deleted.Add(picture.RelId);

            _pictures.Remove(picture);
        }
    }

    IEnumerator<IXlsxPicture> IEnumerable<IXlsxPicture>.GetEnumerator()
    {
        return _pictures.Cast<IXlsxPicture>().GetEnumerator();
    }

    public IEnumerator<XlsxPicture> GetEnumerator()
    {
        return ((IEnumerable<XlsxPicture>)_pictures).GetEnumerator();
    }

    IEnumerator IEnumerable.GetEnumerator()
    {
        return GetEnumerator();
    }

    public IXlsxPicture Picture(string pictureName)
    {
        if (TryGetPicture(pictureName, out IXlsxPicture? p))
            return p!;

        throw new ArgumentOutOfRangeException(nameof(pictureName), $"Picture {pictureName} was not found.");
    }

    public bool TryGetPicture(string pictureName, out IXlsxPicture? picture)
    {
        var match = _pictures.FirstOrDefault(p => p.Name.Equals(pictureName, StringComparison.OrdinalIgnoreCase));
        if (match is not null)
        {
            picture = match;
            return true;
        }
        picture = null;
        return false;
    }

    internal IXlsxPicture Add(Stream stream, string name, int Id)
    {
        var picture = (XlsxPicture)Add(stream);
        picture.SetName(name);
        picture.Id = Id;
        return picture;
    }

    private String GetNextPictureName()
    {
        var pictureNumber = this.Count;
        while (_pictures.Any(p => p.Name == $"Picture {pictureNumber}"))
        {
            pictureNumber++;
        }
        return $"Picture {pictureNumber}";
    }
}
