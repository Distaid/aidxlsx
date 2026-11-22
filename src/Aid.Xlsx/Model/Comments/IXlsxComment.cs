#nullable disable

using Aid.Xlsx.Model.Drawings;
using Aid.Xlsx.Model.RichText;

namespace Aid.Xlsx.Model.Comments;

public interface IXlsxComment : IXlsxFormattedText<IXlsxComment>, IXlsxDrawing<IXlsxComment>
{
    /// <summary>
    /// Gets or sets this comment's author's name
    /// </summary>
    String Author { get; set; }
    /// <summary>
    /// Sets the name of the comment's author
    /// </summary>
    /// <param name="value">Author's name</param>
    IXlsxComment SetAuthor(String value);

    /// <summary>
    /// Adds a bolded line with the author's name
    /// </summary>
    IXlsxRichString AddSignature();

    void Delete();
}
