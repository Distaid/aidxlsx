using System.Reflection;
using Aid.Xlsx.Extensions;

namespace Aid.Xlsx.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = false)]
    public class XlsxColumnAttribute : Attribute
    {
        public String? Header { get; set; }
        public Boolean Ignore { get; set; }
        public Int32 Order { get; set; }

        /// <summary>Excel number format applied by the generated mapper to data cells in this column.</summary>
        public string? Format { get; set; }

        private static XlsxColumnAttribute? GetXLColumnAttribute(MemberInfo mi)
        {
            if (!mi.HasAttribute<XlsxColumnAttribute>()) return null;
            return mi.GetAttributes<XlsxColumnAttribute>().First();
        }

        internal static String? GetHeader(MemberInfo mi)
        {
            var attribute = GetXLColumnAttribute(mi);
            if (attribute == null) return null;
            return String.IsNullOrWhiteSpace(attribute.Header) ? null : attribute.Header;
        }

        internal static Int32 GetOrder(MemberInfo mi)
        {
            var attribute = GetXLColumnAttribute(mi);
            if (attribute == null) return Int32.MaxValue;
            return attribute.Order;
        }

        internal static Boolean IgnoreMember(MemberInfo mi)
        {
            var attribute = GetXLColumnAttribute(mi);
            if (attribute == null) return false;
            return attribute.Ignore;
        }
    }
}
