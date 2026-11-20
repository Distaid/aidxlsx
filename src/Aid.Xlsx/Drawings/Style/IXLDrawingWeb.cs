using System;

namespace Aid.Xlsx
{
    public interface IXLDrawingWeb
    {
        String? AlternateText { get; set; }
        IXLDrawingStyle SetAlternateText(String? value);
    }
}
