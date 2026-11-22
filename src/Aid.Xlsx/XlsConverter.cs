using Aid.Xlsx.Internal.IO;

namespace Aid.Xlsx;

/// <summary>Converts Excel 97–2003 binary workbooks to XLSX.</summary>
/// <remarks>Imports cell values, cached formula results, number formats, dimensions, visibility, and merged cells.</remarks>
public static class XlsConverter
{
    /// <summary>Converts an existing BIFF8 file to an XLSX file at a different path.</summary>
    public static void Convert(string sourcePath, string destinationPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sourcePath);
        ArgumentException.ThrowIfNullOrWhiteSpace(destinationPath);
        if (string.Equals(Path.GetFullPath(sourcePath), Path.GetFullPath(destinationPath),
                OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
            throw new ArgumentException("The source and destination must be different files.", nameof(destinationPath));

        // Finish creating the XLSX before touching any existing destination file.
        using var source = File.OpenRead(sourcePath);
        using var converted = CreateXlsx(source);
        using var destination = File.Create(destinationPath);
        converted.CopyTo(destination);
    }

    /// <summary>Reads a BIFF8 workbook from the current position of a seekable stream and writes XLSX.</summary>
    /// <remarks>Both streams remain open. Use an empty destination stream or position it at the start of the desired output.</remarks>
    public static void Convert(Stream source, Stream destination)
    {
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(destination);
        if (!source.CanRead || !source.CanSeek)
            throw new ArgumentException("The source stream must be readable and seekable.", nameof(source));
        if (!destination.CanWrite || ReferenceEquals(source, destination))
            throw new ArgumentException("The destination must be a different writable stream.", nameof(destination));

        using var converted = CreateXlsx(source);
        converted.CopyTo(destination);
        if (destination.CanSeek)
            destination.SetLength(destination.Position);
    }

    private static MemoryStream CreateXlsx(Stream source)
    {
        using var workbook = new XLWorkbook();
        try
        {
            new Biff8WorkbookReader(CompoundWorkbookStream.Read(source)).ReadInto(workbook);
        }
        catch (Exception ex) when (ex is EndOfStreamException or IndexOutOfRangeException or ArgumentException)
        {
            throw new InvalidDataException("The binary Excel workbook is malformed.", ex);
        }

        var converted = new MemoryStream();
        try
        {
            workbook.SaveAs(converted);
            converted.Position = 0;
            return converted;
        }
        catch
        {
            converted.Dispose();
            throw;
        }
    }
}
