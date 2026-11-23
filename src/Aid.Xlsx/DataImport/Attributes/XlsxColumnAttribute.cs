using System.Reflection;
using Aid.Xlsx.Infrastructure.Extensions;

namespace Aid.Xlsx.DataImport.Attributes;

[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
public class XlsxColumnAttribute : Attribute
{
    public string? Header { get; set; }
    public bool Ignore { get; set; }
    public int Order { get; set; }

    /// <summary>Excel number format applied by the generated mapper to data cells in this column.</summary>
    public string? Format { get; set; }

    private static XlsxColumnAttribute? GetXLColumnAttribute(MemberInfo mi)
    {
        if (!mi.HasAttribute<XlsxColumnAttribute>()) return null;
        return mi.GetAttributes<XlsxColumnAttribute>().First();
    }

    internal static string? GetHeader(MemberInfo mi)
    {
        XlsxColumnAttribute? attribute = GetXLColumnAttribute(mi);
        if (attribute == null) return null;
        return string.IsNullOrWhiteSpace(attribute.Header) ? null : attribute.Header;
    }

    internal static int GetOrder(MemberInfo mi)
    {
        XlsxColumnAttribute? attribute = GetXLColumnAttribute(mi);
        if (attribute == null) return int.MaxValue;
        return attribute.Order;
    }

    internal static bool IgnoreMember(MemberInfo mi)
    {
        XlsxColumnAttribute? attribute = GetXLColumnAttribute(mi);
        if (attribute == null) return false;
        return attribute.Ignore;
    }
}
