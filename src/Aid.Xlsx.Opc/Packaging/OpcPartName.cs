namespace Aid.Xlsx.Opc.Packaging;

internal static class OpcPartName
{
    internal const string RootRelationshipsEntry = "_rels/.rels";
    internal const string ContentTypesEntry = "[Content_Types].xml";

    internal static string Validate(string partName)
    {
        if (partName.Length < 2 || partName[0] != '/')
            throw new InvalidDataException($"Invalid OPC part name '{partName}'.");

        foreach (var segment in partName[1..].Split('/'))
        {
            if (segment.Length == 0 || segment[^1] == '.' || !IsValidSegment(segment))
                throw new InvalidDataException($"Invalid OPC part name '{partName}'.");
        }

        return partName;
    }

    internal static string RelationshipEntry(string? partName)
    {
        if (partName is null)
            return RootRelationshipsEntry;

        var slash = partName.LastIndexOf('/');
        return string.Concat(partName.AsSpan(1, slash), "_rels/", partName.AsSpan(slash + 1), ".rels");
    }

    internal static bool TryGetRelationshipSource(string entryName, out string? source)
    {
        source = null;
        if (entryName.Equals(RootRelationshipsEntry, StringComparison.OrdinalIgnoreCase))
            return true;

        if (!entryName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            return false;

        var slash = entryName.LastIndexOf('/');
        if (slash < 0)
            return false;

        var previousSlash = entryName.LastIndexOf('/', slash - 1);
        if (!entryName.AsSpan(previousSlash + 1, slash - previousSlash - 1).Equals("_rels", StringComparison.OrdinalIgnoreCase))
            return false;

        source = string.Concat("/", entryName.AsSpan(0, previousSlash + 1), entryName.AsSpan(slash + 1, entryName.Length - slash - 6));
        return true;
    }

    internal static string ResolveTarget(string? source, string target)
    {
        var fragment = target.IndexOf('#');
        var path = fragment < 0 ? target : target[..fragment];
        if (fragment >= 0 && !IsValidFragment(target.AsSpan(fragment + 1)))
            throw new InvalidDataException($"Invalid internal OPC target fragment '{target}'.");
        if (path.Length == 0 && source is not null && fragment == 0)
            return source;

        var firstSlash = path.IndexOf('/');
        var firstSegment = firstSlash < 0 ? path.AsSpan() : path.AsSpan(0, firstSlash);
        if (path.Length == 0 || path.EndsWith('/') || firstSegment.Contains(':'))
            throw new InvalidDataException($"Invalid internal OPC target '{target}'.");

        var segments = new List<string>();
        if (path[0] == '/')
            path = path[1..];
        else if (source is not null)
        {
            var slash = source.LastIndexOf('/');
            if (slash > 0)
                segments.AddRange(source[1..slash].Split('/'));
        }

        foreach (var segment in path.Split('/'))
        {
            switch (segment)
            {
                case ".":
                    break;
                case ".." when segments.Count > 0:
                    segments.RemoveAt(segments.Count - 1);
                    break;
                default:
                    if (segment.Length == 0 || segment[^1] == '.' || !IsValidSegment(segment))
                        throw new InvalidDataException($"Invalid internal OPC target '{target}'.");
                    segments.Add(segment);
                    break;
            }
        }

        return Validate("/" + string.Join('/', segments));
    }

    private static bool IsValidSegment(string segment)
    {
        for (var i = 0; i < segment.Length; i++)
        {
            var c = segment[i];
            if (IsUnreserved(c) || "!$&'()*+,;=:@".Contains(c))
                continue;

            if (c != '%' || i + 2 >= segment.Length || !Uri.IsHexDigit(segment[i + 1]) || !Uri.IsHexDigit(segment[i + 2]))
                return false;

            var decoded = (char)((Uri.FromHex(segment[i + 1]) << 4) | Uri.FromHex(segment[i + 2]));
            if (decoded is '/' or '\\' or < ' ' or '\u007F' || IsUnreserved(decoded))
                return false;
            i += 2;
        }

        return true;
    }

    private static bool IsUnreserved(char c) => char.IsAsciiLetterOrDigit(c) || c is '-' or '.' or '_' or '~';

    private static bool IsValidFragment(ReadOnlySpan<char> fragment)
    {
        for (var i = 0; i < fragment.Length; i++)
        {
            var c = fragment[i];
            if (IsUnreserved(c) || "!$&'()*+,;=:@/?".Contains(c))
                continue;

            if (c != '%' || i + 2 >= fragment.Length || !Uri.IsHexDigit(fragment[i + 1]) || !Uri.IsHexDigit(fragment[i + 2]))
                return false;
            i += 2;
        }

        return true;
    }
}
