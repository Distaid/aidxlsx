using System.Buffers.Binary;
using Aid.Xlsx.CalcEngine;
using Aid.Xlsx.Cells;
using Aid.Xlsx.Styles.Formatting;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;
using Array = System.Array;

namespace Aid.Xlsx.Internal.IO;

internal sealed partial class Biff8WorkbookReader(byte[] data)
{
    private const ushort Bof = 0x0809;
    private const ushort Eof = 0x000A;
    private const ushort Continue = 0x003C;
    private const ushort FormulaString = 0x0207;
    private const int MaxSheets = 1024;
    private const int MaxSharedStrings = 1_000_000;
    private const int MaxMaterializedCells = 5_000_000;
    private readonly List<SheetInfo> _sheets = [];
    private string[] _strings = [];
    private int _materializedCells;

    public void ReadInto(XlsxWorkbook workbook)
    {
        var offset = 0;
        ReadBof(ref offset, 0x0005);
        while (true)
        {
            var record = ReadRecord(ref offset, out var type);
            if (type == Eof)
                break;

            switch (type)
            {
                case 0x002F: // FILEPASS
                    throw new NotSupportedException("Encrypted .xls workbooks are not supported.");
                case 0x0022: // DATEMODE
                    workbook.Use1904DateSystem = U16(record) != 0;
                    break;
                case 0x0085: // BOUNDSHEET8
                    ReadSheetInfo(record);
                    break;
                case 0x00E0: // XF
                    ReadCellFormat(record);
                    break;
                case 0x0031: // FONT
                    ReadFont(record);
                    break;
                case 0x0092: // PALETTE
                    ReadPalette(record);
                    break;
                case 0x041E: // FORMAT
                    _customFormats[U16(record)] = ReadUnicodeString(record[2..], ref offset);
                    break;
                case 0x00FC: // SST
                    ReadSharedStrings(record, ref offset);
                    break;
            }
        }

        if (_sheets.Count == 0)
            throw new InvalidDataException("The .xls workbook contains no worksheets.");

        _cellFormats = new XlsxCellFormatValue?[_rawCellFormats.Count];
        var boundaries = _sheets.Select(sheet => sheet.Offset).Order().ToArray();
        for (var i = 0; i < boundaries.Length; i++)
        {
            if (boundaries[i] < offset || boundaries[i] >= data.Length || (i > 0 && boundaries[i] == boundaries[i - 1]))
                throw new InvalidDataException("Invalid BIFF worksheet offset.");
        }

        foreach (var sheet in _sheets)
        {
            var index = Array.BinarySearch(boundaries, sheet.Offset);
            var end = index + 1 < boundaries.Length ? boundaries[index + 1] : data.Length;
            var worksheet = (XlsxWorksheet)workbook.Worksheets.Add(sheet.Name);
            worksheet.Visibility = sheet.Visibility;
            ReadWorksheet(worksheet, sheet.Offset, end);
        }
    }

    private void ReadSheetInfo(ReadOnlySpan<byte> record)
    {
        if (_sheets.Count >= MaxSheets)
            throw new InvalidDataException("The .xls workbook exceeds the worksheet limit.");
        if (record[5] != 0)
            throw new NotSupportedException("Chart, macro, and dialog sheets in .xls workbooks are not supported.");

        var visibility = record[4] switch
        {
            0 => XlsxWorksheetVisibility.Visible,
            1 => XlsxWorksheetVisibility.Hidden,
            2 => XlsxWorksheetVisibility.VeryHidden,
            _ => throw new InvalidDataException("Invalid BIFF sheet visibility.")
        };
        var offset = U32(record);
        var length = record[6];
        if (offset > int.MaxValue || length is 0 or > 31)
            throw new InvalidDataException("Invalid BIFF sheet header.");

        var name = ReadCharacters(record[8..], length, (record[7] & 1) != 0);
        _sheets.Add(new SheetInfo((int)offset, name, visibility));
    }

    private void ReadWorksheet(XlsxWorksheet sheet, int offset, int end)
    {
        ReadBof(ref offset, 0x0010);
        while (offset < end)
        {
            var record = ReadRecord(ref offset, out var type);
            if (offset > end)
                break;

            switch (type)
            {
                case Eof:
                    return;
                case 0x0203: // NUMBER
                    SetCell(sheet, record, BinaryPrimitives.ReadDoubleLittleEndian(record[6..]));
                    break;
                case 0x027E: // RK
                    SetCell(sheet, record, ReadRk(record[6..]));
                    break;
                case 0x00BD: // MULRK
                    ReadMultipleCells(sheet, record, hasValues: true);
                    break;
                case 0x00BE: // MULBLANK
                    ReadMultipleCells(sheet, record, hasValues: false);
                    break;
                case 0x0201: // BLANK
                    SetCell(sheet, record, Blank.Value);
                    break;
                case 0x00FD: // LABELSST
                    var index = U32(record[6..]);
                    if (index >= _strings.Length)
                        throw new InvalidDataException("Invalid BIFF shared string index.");
                    SetCell(sheet, record, _strings[index]);
                    break;
                case 0x0204: // LABEL
                case 0x00D6: // RSTRING
                    SetCell(sheet, record, ReadUnicodeString(record[6..], ref offset));
                    break;
                case 0x0205: // BOOLERR
                    SetCell(sheet, record, record[7] == 0 ? (XlsxCellValue)(record[6] != 0) : ReadError(record[6]));
                    break;
                case 0x0006: // FORMULA
                    ReadFormulaCell(sheet, record, ref offset);
                    break;
                case 0x0208: // ROW
                    ReadRowInfo(sheet, record);
                    break;
                case 0x007D: // COLINFO
                    ReadColumnInfo(sheet, record);
                    break;
                case 0x00E5: // MERGEDCELLS
                    ReadMerges(sheet, record);
                    break;
            }
        }
        throw new InvalidDataException("Missing BIFF worksheet end record.");
    }

    private void ReadSharedStrings(ReadOnlySpan<byte> record, ref int offset)
    {
        var count = U32(record[4..]);
        if (count > MaxSharedStrings || count > (data.Length - offset + (long)record.Length) / 3 || count > U32(record))
            throw new InvalidDataException("Invalid or excessively large BIFF shared string table.");

        var segments = new StringSegments(data, record[8..], offset);
        var strings = new string[count];
        for (var i = 0; i < strings.Length; i++)
            strings[i] = segments.ReadString();
        _strings = strings;
        offset = segments.Offset;
    }

    private string ReadUnicodeString(ReadOnlySpan<byte> record, ref int offset)
    {
        var segments = new StringSegments(data, record, offset);
        var value = segments.ReadString();
        offset = segments.Offset;
        return value;
    }

    private void ReadFormulaCell(XlsxWorksheet sheet, ReadOnlySpan<byte> record, ref int offset)
    {
        if (record.Length < 22 || U16(record[20..]) > record.Length - 22)
            throw new InvalidDataException("Invalid BIFF formula length.");

        var value = ReadFormulaResult(record[6..], ref offset);
        var cell = SetCell(sheet, record, value);
        var formula = Biff8Formula.Decode(record.Slice(22, U16(record[20..])));
        if (formula is not null)
            cell.Formula = XlsxCellFormula.NormalA1(formula);
    }

    private XlsxCellValue ReadFormulaResult(ReadOnlySpan<byte> record, ref int offset)
    {
        if (U16(record[6..]) != 0xFFFF)
            return BinaryPrimitives.ReadDoubleLittleEndian(record);

        switch (record[0])
        {
            case 0:
                // Array, shared-formula, or data-table metadata may precede the STRING record.
                while (true)
                {
                    var text = ReadRecord(ref offset, out var type);
                    if (type == FormulaString)
                        return ReadUnicodeString(text, ref offset);
                    if (type is not (0x0221 or 0x04BC or 0x0236))
                        throw new InvalidDataException("Missing BIFF formula string result.");
                }
            case 1:
                return record[2] != 0;
            case 2:
                return ReadError(record[2]);
            case 3:
                return string.Empty;
            default:
                throw new InvalidDataException("Invalid BIFF formula result.");
        }
    }

    private void ReadMultipleCells(XlsxWorksheet sheet, ReadOnlySpan<byte> record, bool hasValues)
    {
        var row = U16(record);
        var first = U16(record[2..]);
        var last = U16(record[^2..]);
        var size = hasValues ? 6 : 2;
        if (last < first || last > 255 || record.Length != 6 + (last - first + 1) * size)
            throw new InvalidDataException("Invalid BIFF multiple-cell record.");

        for (var column = (int)first; column <= last; column++)
        {
            var field = record[(4 + (column - first) * size)..];
            SetCell(sheet, row, column, U16(field), hasValues ? ReadRk(field[2..]) : (XlsxCellValue)Blank.Value);
        }
    }

    private XlsxCell SetCell(XlsxWorksheet sheet, ReadOnlySpan<byte> record, XlsxCellValue value)
        => SetCell(sheet, U16(record), U16(record[2..]), U16(record[4..]), value);

    private XlsxCell SetCell(XlsxWorksheet sheet, int row, int column, int xf, XlsxCellValue value)
    {
        if (column > 255 || xf >= _cellFormats.Length || ++_materializedCells > MaxMaterializedCells)
            throw new InvalidDataException("Invalid BIFF cell index or workbook cell limit exceeded.");

        var cell = sheet.Cell(row + 1, column + 1);
        cell.Value = value;
        cell.FormatValue = _cellFormats[xf] ??= BuildCellFormat(sheet.Workbook.Styles, _rawCellFormats[xf]);
        return cell;
    }

    private static void ReadRowInfo(IXlsxWorksheet sheet, ReadOnlySpan<byte> record)
    {
        var flags = U16(record[12..]);
        if ((flags & 0x60) == 0)
            return;

        var row = sheet.Row(U16(record) + 1);
        if ((flags & 0x40) != 0)
            row.Height = (U16(record[6..]) & 0x7FFF) / 20d;
        if ((flags & 0x20) != 0)
            row.Hide();
    }

    private static void ReadColumnInfo(IXlsxWorksheet sheet, ReadOnlySpan<byte> record)
    {
        var first = U16(record);
        var last = Math.Min((int)U16(record[2..]), 255);
        if (first > last)
            return;

        var width = U16(record[4..]) / 256d;
        var hidden = (U16(record[8..]) & 1) != 0;
        for (var index = (int)first; index <= last; index++)
        {
            var column = sheet.Column(index + 1);
            column.Width = width;
            if (hidden)
                column.Hide();
        }
    }

    private void ReadMerges(IXlsxWorksheet sheet, ReadOnlySpan<byte> record)
    {
        var count = U16(record);
        if (record.Length != 2 + count * 8)
            throw new InvalidDataException("Invalid BIFF merged cell record length.");

        for (var index = 0; index < count; index++)
        {
            var field = record[(2 + index * 8)..];
            var firstRow = U16(field);
            var lastRow = U16(field[2..]);
            var firstColumn = U16(field[4..]);
            var lastColumn = U16(field[6..]);
            if (firstRow > lastRow || firstColumn > lastColumn || lastColumn > 255)
                throw new InvalidDataException("Invalid BIFF merged cell range.");
            _materializedCells += (lastRow - firstRow + 1) * (lastColumn - firstColumn + 1);
            if (_materializedCells > MaxMaterializedCells)
                throw new InvalidDataException("The .xls workbook exceeds the cell limit.");

            sheet.Range(firstRow + 1, firstColumn + 1, lastRow + 1, lastColumn + 1).Merge();
        }
    }

    private void ReadBof(ref int offset, ushort expectedSubstream)
    {
        var record = ReadRecord(ref offset, out var type);
        if (type != Bof || U16(record) != 0x0600)
            throw new NotSupportedException("Only Excel 97–2003 (BIFF8) .xls workbooks are supported.");
        if (U16(record[2..]) != expectedSubstream)
            throw new InvalidDataException("Unexpected BIFF substream type.");
    }

    private ReadOnlySpan<byte> ReadRecord(scoped ref int offset, out ushort type) => ReadRecord(data, ref offset, out type);

    private static ReadOnlySpan<byte> ReadRecord(byte[] buffer, scoped ref int offset, out ushort type)
    {
        if (buffer.Length - offset < 4)
            throw new InvalidDataException("Truncated BIFF record header.");

        var header = buffer.AsSpan(offset);
        type = U16(header);
        var length = U16(header[2..]);
        if (length > 8224 || length > header.Length - 4)
            throw new InvalidDataException("Invalid or truncated BIFF record.");

        offset += 4 + length;
        return header.Slice(4, length);
    }

    private static string ReadCharacters(ReadOnlySpan<byte> record, int count, bool wide)
    {
        var chars = new char[count];
        for (var i = 0; i < chars.Length; i++)
            chars[i] = wide ? (char)U16(record[(2 * i)..]) : (char)record[i];
        return new string(chars);
    }

    private static double ReadRk(ReadOnlySpan<byte> data)
    {
        var bits = BinaryPrimitives.ReadInt32LittleEndian(data);
        var number = (bits & 2) != 0 ? bits >> 2 : BitConverter.Int64BitsToDouble((long)(bits & ~3) << 32);
        return (bits & 1) != 0 ? number / 100 : number;
    }

    private static XlsxError ReadError(byte code) => code switch
    {
        0x00 => XlsxError.NullValue,
        0x07 => XlsxError.DivisionByZero,
        0x0F => XlsxError.IncompatibleValue,
        0x17 => XlsxError.CellReference,
        0x1D => XlsxError.NameNotRecognized,
        0x24 => XlsxError.NumberInvalid,
        0x2A => XlsxError.NoValueAvailable,
        _ => throw new InvalidDataException("Invalid BIFF cell error code.")
    };

    private static ushort U16(ReadOnlySpan<byte> bytes) => BinaryPrimitives.ReadUInt16LittleEndian(bytes);

    private static uint U32(ReadOnlySpan<byte> bytes) => BinaryPrimitives.ReadUInt32LittleEndian(bytes);

    private sealed record SheetInfo(int Offset, string Name, XlsxWorksheetVisibility Visibility);

    private ref struct StringSegments(byte[] buffer, ReadOnlySpan<byte> initial, int offset)
    {
        private ReadOnlySpan<byte> _available = initial;
        private int _offset = offset;

        public readonly int Offset => _offset;

        public string ReadString()
        {
            var length = ReadU16();
            var flags = ReadByte();
            var runs = (flags & 8) != 0 ? ReadU16() : 0;
            var extensionBytes = (flags & 4) != 0 ? ReadU32() : 0;
            if (extensionBytes > buffer.Length)
                throw new InvalidDataException("Invalid BIFF string extension size.");

            var wide = (flags & 1) != 0;
            var chars = new char[length];
            for (var i = 0; i < chars.Length; i++)
            {
                if (_available.IsEmpty)
                {
                    ReadContinuation();
                    wide = (ReadByte() & 1) != 0;
                }
                if (_available.Length < (wide ? 2 : 1))
                    throw new InvalidDataException("A BIFF continuation splits a character.");

                chars[i] = wide ? (char)ReadU16() : (char)ReadByte();
            }

            Skip(runs * 4 + (int)extensionBytes);
            return new string(chars);
        }

        private uint ReadU32() => (uint)ReadU16() | ((uint)ReadU16() << 16);

        private ushort ReadU16() => (ushort)(ReadByte() | (ReadByte() << 8));

        private byte ReadByte()
        {
            if (_available.IsEmpty)
                ReadContinuation();
            var value = _available[0];
            _available = _available[1..];
            return value;
        }

        private void Skip(int count)
        {
            while (count > 0)
            {
                if (_available.IsEmpty)
                    ReadContinuation();
                var consume = Math.Min(count, _available.Length);
                _available = _available[consume..];
                count -= consume;
            }
        }

        private void ReadContinuation()
        {
            var bytes = ReadRecord(buffer, ref _offset, out var type);
            if (type != Continue || bytes.IsEmpty)
                throw new InvalidDataException("Missing BIFF string continuation record.");
            _available = bytes;
        }
    }
}
