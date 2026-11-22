using System.Diagnostics.CodeAnalysis;
using Aid.Xlsx.Ranges;

namespace Aid.Xlsx.DefinedNames
{
    public interface IXlsxDefinedNames : IEnumerable<IXlsxDefinedName>
    {
        /// <inheritdoc cref="DefinedName"/>
        [Obsolete($"Use {nameof(DefinedName)} instead.")]
        IXlsxDefinedName NamedRange(String name);

        /// <summary>
        /// Gets the specified defined name.
        /// </summary>
        /// <param name="name">Name identifier.</param>
        /// <exception cref="KeyNotFoundException">Name wasn't found.</exception>
        IXlsxDefinedName DefinedName(String name);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="rangeAddress">The range address to add.</param>
        /// <exception cref="ArgumentException">The name or address is invalid.</exception>
        IXlsxDefinedName Add(String name, String rangeAddress);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="range">The range to add.</param>
        /// <exception cref="ArgumentException">The name is invalid.</exception>
        IXlsxDefinedName Add(String name, IXlsxRange range);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="ranges">The ranges to add.</param>
        /// <exception cref="ArgumentException">The name is invalid.</exception>
        IXlsxDefinedName Add(String name, IXlsxRanges ranges);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="rangeAddress">The range address to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        /// <exception cref="ArgumentException">The range name or address is invalid.</exception>
        IXlsxDefinedName Add(String name, String rangeAddress, String? comment);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="range">The range to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        /// <exception cref="ArgumentException">The range name is invalid.</exception>
        IXlsxDefinedName Add(String name, IXlsxRange range, String? comment);

        /// <summary>
        /// Adds a new defined name.
        /// </summary>
        /// <param name="name">Name identifier to add.</param>
        /// <param name="ranges">The ranges to add.</param>
        /// <param name="comment">The comment for the new named range.</param>
        /// <exception cref="ArgumentException">The range name is invalid.</exception>
        IXlsxDefinedName Add(String name, IXlsxRanges ranges, String? comment);

        /// <summary>
        /// Deletes the specified defined name.  Deleting defined name doesn't delete referenced
        /// cells.
        /// </summary>
        /// <param name="name">Name identifier to delete.</param>
        void Delete(String name);

        /// <summary>
        /// Deletes the specified defined name's index. Deleting defined name doesn't delete
        /// referenced cells.
        /// </summary>
        /// <param name="index">Index of the defined name to delete.</param>
        /// <exception cref="ArgumentOutOfRangeException">The index is outside of named ranges array.</exception>
        void Delete(Int32 index);

        /// <summary>
        /// Deletes all defined names of this collection, i.e. a workbook or a sheet. Deleting
        /// defined name doesn't delete referenced cells.
        /// </summary>
        void DeleteAll();

        Boolean TryGetValue(String name, [NotNullWhen(true)] out IXlsxDefinedName? range);

        Boolean Contains(String name);

        /// <summary>
        /// Returns a subset of defined names that do not have invalid references.
        /// </summary>
        IEnumerable<IXlsxDefinedName> ValidNamedRanges();

        /// <summary>
        /// Returns a subset of defined names that do have invalid references.
        /// </summary>
        IEnumerable<IXlsxDefinedName> InvalidNamedRanges();
    }
}
