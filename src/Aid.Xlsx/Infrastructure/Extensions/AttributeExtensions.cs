#nullable disable

// Keep this file CodeMaid organised and cleaned
using System.Reflection;

namespace Aid.Xlsx.Extensions
{
    internal static class AttributeExtensions
    {
        public static TAttribute[] GetAttributes<TAttribute>(
            this MemberInfo member)
            where TAttribute : Attribute
        {
            var attributes = member.GetCustomAttributes(typeof(TAttribute), true);

            return (TAttribute[])attributes;
        }

        public static bool HasAttribute<TAttribute>(
            this MemberInfo member)
            where TAttribute : Attribute
        {
            return GetAttributes<TAttribute>(member).Any();
        }
    }
}
