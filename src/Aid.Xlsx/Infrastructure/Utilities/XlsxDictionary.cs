#nullable disable

using Aid.Xlsx.Extensions;

namespace Aid.Xlsx.Utils
{
    public class XlsxDictionary<T> : Dictionary<Int32, T>
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
            foreach (var (key, value) in this)
                copy.Add(key, value);

            return copy;
        }
    }
}
