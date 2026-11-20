using System;
using System.Collections.Generic;

namespace Aid.Xlsx
{
    public interface IXLConditionalFormats : IEnumerable<IXLConditionalFormat>
    {
        void Add(IXLConditionalFormat conditionalFormat);

        void RemoveAll();

        void Remove(Predicate<IXLConditionalFormat> predicate);
    }
}
