namespace Aid.Xlsx.ConditionalFormats
{
    public interface IXlsxConditionalFormats : IEnumerable<IXlsxConditionalFormat>
    {
        void Add(IXlsxConditionalFormat conditionalFormat);

        void RemoveAll();

        void Remove(Predicate<IXlsxConditionalFormat> predicate);
    }
}
