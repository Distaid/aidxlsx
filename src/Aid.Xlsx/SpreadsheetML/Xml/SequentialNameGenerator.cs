using System.Globalization;
using Aid.Xlsx.Infrastructure.Utilities;

namespace Aid.Xlsx.SpreadsheetML.Xml;

internal class SequentialNameGenerator
{
    private readonly string _prefix;
    private int _nextNumber;

    internal SequentialNameGenerator(string prefix, int nextNumber)
    {
        _prefix = prefix;
        _nextNumber = nextNumber;
    }

    internal void AddName(string name)
    {
        if (!name.StartsWith(_prefix))
            return;

        if (!int.TryParse(name[_prefix.Length..], NumberStyles.None, XlsxHelper.ParseCulture, out var styleNumber))
            return;

        _nextNumber = Math.Max(styleNumber + 1, _nextNumber);
    }

    internal string NextUnusedStyleName()
    {
        return $"{_prefix}{_nextNumber++}";
    }
}
