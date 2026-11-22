// Keep this file CodeMaid organised and cleaned
using System.Collections;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using Aid.Xlsx.DataImport.Attributes;
using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Grid.Cells;

namespace Aid.Xlsx.DataImport.RowReaders;

internal class ObjectReader : IInsertDataReader
{
    private const BindingFlags MemberBindingFlags = BindingFlags.Public
                                              | BindingFlags.Instance
                                              | BindingFlags.Static;

    private readonly IEnumerable<object> _data;
    private readonly MemberInfo[] _members;
    private readonly bool[] _staticMembers;

    [RequiresUnreferencedCode("Objects are read as rows with reflection over their public fields and properties. Use a DataTable or explicit XlsxCellValue rows to stay trim-compatible.")]
    public ObjectReader(IEnumerable data)
        : this(data, data.GetItemType()!)
    {
    }

    /// <summary>
    /// Create a reader for objects of a known item type. The type is passed explicitly, because the
    /// sequence can be a non-generic projection (e.g. <c>List&lt;object&gt;</c> used to split mixed
    /// data by runtime type) where <see cref="EnumerableExtensions.GetItemType"/> would return
    /// <see cref="object"/> instead of the real item type.
    /// </summary>
    [RequiresUnreferencedCode("Objects are read as rows with reflection over their public fields and properties. Use a DataTable or explicit XlsxCellValue rows to stay trim-compatible.")]
    internal ObjectReader(IEnumerable data, Type itemType)
    {
        _data = data.Cast<object>();

        if (itemType.IsNullableType())
            itemType = itemType.GetUnderlyingType();

        _members = itemType.GetFields(MemberBindingFlags).Cast<MemberInfo>()
            .Concat(itemType.GetProperties(MemberBindingFlags).Where(pi => !pi.GetIndexParameters().Any()))
            .Where(mi => !XlsxColumnAttribute.IgnoreMember(mi))
            .OrderBy(XlsxColumnAttribute.GetOrder)
            .ToArray();

        _staticMembers = _members.Select(ReflectionExtensions.IsStatic).ToArray();
    }

    public IEnumerable<IEnumerable<XlsxCellValue>> GetRecords()
    {
        return _data.Select(item => GetItemData(item).Select(XlsxCellValue.FromInsertedObject));
    }

    public int GetPropertiesCount()
    {
        return _members.Length;
    }

    public string? GetPropertyName(int propertyIndex)
    {
        if (propertyIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(propertyIndex), "Property index must be non-negative");

        if (propertyIndex >= GetPropertiesCount())
            throw new ArgumentOutOfRangeException($"{propertyIndex} exceeds the number of the object properties");

        var memberInfo = _members[propertyIndex];
        var fieldName = XlsxColumnAttribute.GetHeader(memberInfo);
        if (String.IsNullOrWhiteSpace(fieldName))
            fieldName = memberInfo.Name;

        return fieldName;
    }

    private IEnumerable<object?> GetItemData(object item)
    {
        for (int i = 0; i < _members.Length; i++)
        {
            if (item == null)
            {
                yield return null;
                continue;
            }

            var memberInfo = _members[i];
            switch (memberInfo)
            {
                case PropertyInfo propertyInfo when _staticMembers[i]:
                    yield return propertyInfo.GetValue(null, null);
                    break;

                case PropertyInfo propertyInfo:
                    yield return propertyInfo.GetValue(item, null);
                    break;

                case FieldInfo fieldInfo when _staticMembers[i]:
                    yield return fieldInfo.GetValue(null);
                    break;

                case FieldInfo fieldInfo:
                    yield return fieldInfo.GetValue(item);
                    break;
            }
        }
    }
}
