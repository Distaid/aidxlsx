#nullable disable

// Keep this file CodeMaid organised and cleaned
namespace Aid.Xlsx.Infrastructure.Extensions;

internal static class DictionaryExtensions
{
    public static void RemoveAll<TKey, TValue>(this Dictionary<TKey, TValue> dic,
        Func<TValue, bool> predicate)
    {
        var keys = dic.Keys.Where(k => predicate(dic[k])).ToList();
        foreach (TKey key in keys)
        {
            dic.Remove(key);
        }
    }
}
