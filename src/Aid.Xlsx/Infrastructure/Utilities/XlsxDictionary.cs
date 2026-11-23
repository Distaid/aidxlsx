#nullable disable

using Aid.Xlsx.Infrastructure.Extensions;

namespace Aid.Xlsx.Infrastructure.Utilities;

public class XlsxDictionary<T> : Dictionary<int, T>
{
    public XlsxDictionary()
    {

    }
    public XlsxDictionary(XlsxDictionary<T> other)
    {
        other.Values.ForEach(Add);
    }

    public void Initialize(T value)
    {
        if (Count > 0)
            Clear();

        Add(value);
    }

    public void Add(T value)
    {
        Add(Count + 1, value);
    }

    internal XlsxDictionary<T> CopyDictionary()
    {
        var copy = new XlsxDictionary<T>();
        foreach ((var key, T value) in this)
            copy.Add(key, value);

        return copy;
    }
}
