namespace Aid.Xlsx.Tests.TestSupport;

// TODO: Replace with EnumPolyfill once Polyfill is updated
internal class EnumPolyfill
{
    public static T[] GetValues<T>()
        where T : struct, Enum
    {
        return Enum.GetValues(typeof(T)).Cast<T>().ToArray();
    }
}
