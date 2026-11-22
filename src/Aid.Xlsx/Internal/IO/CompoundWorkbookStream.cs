using System.Buffers.Binary;
using System.Text;

namespace Aid.Xlsx.Internal.IO;

/// <summary>Extracts the root Workbook or Book stream from a compound file (MS-CFB).</summary>
internal sealed class CompoundWorkbookStream
{
    private const uint EndOfChain = 0xFFFFFFFE;
    private const uint Free = 0xFFFFFFFF;
    private const int MaxWorkbookBytes = 128 * 1024 * 1024;
    private const int MiniSectorSize = 64;
    private const int MiniStreamCutoff = 4096;
    private readonly Stream _stream;
    private readonly long _origin;
    private readonly int _sectorSize;
    private readonly int _sectorCount;
    private readonly uint[] _fat;

    private CompoundWorkbookStream(Stream stream, long origin, ReadOnlySpan<byte> header)
    {
        _stream = stream;
        _origin = origin;
        var version = U16(header[26..]);
        var shift = U16(header[30..]);
        if ((version, shift) is not ((3, 9) or (4, 12)) || U16(header[28..]) != 0xFFFE ||
            U16(header[32..]) != 6 || U32(header[56..]) != MiniStreamCutoff)
            throw Invalid("Invalid compound file header.");

        _sectorSize = 1 << shift;
        var length = stream.Length - origin;
        if (length < 2L * _sectorSize || length % _sectorSize != 0 || length > 2L * MaxWorkbookBytes)
            throw Invalid("Invalid or excessively large compound file.");

        _sectorCount = (int)(length / _sectorSize) - 1;
        _fat = ReadFat(header);
    }

    public static byte[] Read(Stream stream)
    {
        var origin = stream.Position;
        Span<byte> signature = stackalloc byte[8];
        stream.ReadExactly(signature);
        stream.Position = origin;
        if (!signature.SequenceEqual(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }))
        {
            var length = stream.Length - origin;
            if (U16(signature) != 0x0809 || length > MaxWorkbookBytes)
                throw Invalid("Expected a binary Excel workbook.");

            var raw = new byte[(int)length];
            stream.ReadExactly(raw);
            return raw;
        }

        Span<byte> header = stackalloc byte[512];
        stream.ReadExactly(header);
        return new CompoundWorkbookStream(stream, origin, header).ReadWorkbook(header);
    }

    private uint[] ReadFat(ReadOnlySpan<byte> header)
    {
        var fatCount = U32(header[44..]);
        var difatCount = U32(header[72..]);
        var entriesPerSector = _sectorSize / sizeof(uint);
        if (fatCount == 0 || fatCount > (_sectorCount + entriesPerSector - 1) / entriesPerSector ||
            (long)fatCount * entriesPerSector < _sectorCount || difatCount > _sectorCount)
            throw Invalid("Invalid compound file allocation table size.");

        var fat = new uint[fatCount * entriesPerSector];
        var data = new byte[_sectorSize];
        var seen = new HashSet<uint>();
        var offset = 0;
        for (var i = 76; i < header.Length && offset < fat.Length; i += sizeof(uint))
            ReadFatSector(U32(header[i..]));

        var difat = U32(header[68..]);
        var difatSector = new byte[_sectorSize];
        for (var i = 0; i < difatCount; i++)
        {
            AddSector(difat);
            ReadSector(difat, difatSector);
            for (var j = 0; j < _sectorSize - sizeof(uint) && offset < fat.Length; j += sizeof(uint))
                ReadFatSector(U32(difatSector.AsSpan(j)));
            difat = U32(difatSector.AsSpan(_sectorSize - sizeof(uint)));
        }

        if (offset != fat.Length || difat != EndOfChain)
            throw Invalid("Incomplete compound file allocation table.");
        return fat;

        void ReadFatSector(uint sector)
        {
            if (sector == Free)
                return;

            AddSector(sector);
            ReadSector(sector, data);
            for (var j = 0; j < data.Length; j += sizeof(uint))
                fat[offset++] = U32(data.AsSpan(j));
        }

        void AddSector(uint sector)
        {
            if (sector >= _sectorCount || !seen.Add(sector))
                throw Invalid("Invalid or repeated allocation table sector.");
        }
    }

    private byte[] ReadWorkbook(ReadOnlySpan<byte> header)
    {
        using var directory = new MemoryStream();
        var sector = U32(header[48..]);
        var seen = new HashSet<uint>();
        var buffer = new byte[_sectorSize];
        while (sector != EndOfChain)
        {
            if (directory.Length == 2 * 1024 * 1024 || !seen.Add(sector))
                throw Invalid("Invalid compound file directory chain.");

            ReadSector(sector, buffer);
            directory.Write(buffer);
            sector = _fat[sector];
        }

        var entries = directory.GetBuffer().AsSpan(0, (int)directory.Length);
        if (entries.Length < 128 || entries[66] != 5)
            throw Invalid("Missing compound file root directory entry.");

        var root = entries[..128];
        var pending = new Stack<uint>();
        pending.Push(U32(root[76..]));
        seen.Clear();
        while (pending.TryPop(out var entryIndex))
        {
            if (entryIndex == Free)
                continue;
            if (entryIndex == 0 || entryIndex >= entries.Length / 128 || !seen.Add(entryIndex))
                throw Invalid("Invalid compound file directory tree.");

            var entry = entries.Slice((int)entryIndex * 128, 128);
            var nameBytes = U16(entry[64..]);
            if (nameBytes is < 2 or > 64 || (nameBytes & 1) != 0)
                throw Invalid("Invalid compound file directory name.");

            var name = Encoding.Unicode.GetString(entry[..(nameBytes - 2)]);
            if (entry[66] == 2 && (name.Equals("Workbook", StringComparison.OrdinalIgnoreCase) ||
                                   name.Equals("Book", StringComparison.OrdinalIgnoreCase)))
            {
                var size = StreamSize(entry, header);
                var start = U32(entry[116..]);
                return size >= MiniStreamCutoff ? ReadRegularChain(start, size) : ReadMiniChain(root, header, start, size);
            }

            pending.Push(U32(entry[68..]));
            pending.Push(U32(entry[72..]));
        }
        throw Invalid("The compound file has no root Workbook or Book stream.");
    }

    private byte[] ReadRegularChain(uint sector, int size)
    {
        var output = new byte[size];
        var buffer = new byte[_sectorSize];
        var seen = new HashSet<uint>();
        for (var offset = 0; offset < size; offset += _sectorSize)
        {
            if (!seen.Add(sector))
                throw Invalid("Repeated compound file stream sector.");

            ReadSector(sector, buffer);
            buffer.AsSpan(0, Math.Min(_sectorSize, size - offset)).CopyTo(output.AsSpan(offset));
            sector = _fat[sector];
        }
        if (sector != EndOfChain)
            throw Invalid("Invalid compound file stream length.");

        return output;
    }

    private byte[] ReadMiniChain(ReadOnlySpan<byte> root, ReadOnlySpan<byte> header, uint sector, int size)
    {
        var miniFatSectors = U32(header[64..]);
        if (miniFatSectors == 0 || miniFatSectors > _sectorCount)
            throw Invalid("Invalid compound file mini allocation table size.");

        var miniStream = ReadRegularChain(U32(root[116..]), StreamSize(root, header));
        var miniFat = ReadRegularChain(U32(header[60..]), (int)miniFatSectors * _sectorSize);
        var output = new byte[size];
        var seen = new HashSet<uint>();
        for (var offset = 0; offset < size; offset += MiniSectorSize)
        {
            var count = Math.Min(MiniSectorSize, size - offset);
            if (sector >= miniFat.Length / sizeof(uint) || (long)sector * MiniSectorSize + count > miniStream.Length || !seen.Add(sector))
                throw Invalid("Invalid compound file mini stream chain.");

            miniStream.AsSpan((int)sector * MiniSectorSize, count).CopyTo(output.AsSpan(offset));
            sector = U32(miniFat.AsSpan((int)sector * sizeof(uint)));
        }
        if (sector != EndOfChain)
            throw Invalid("Invalid compound file mini stream length.");

        return output;
    }

    private void ReadSector(uint sector, Span<byte> buffer)
    {
        if (sector >= _sectorCount)
            throw Invalid("Compound file sector is outside the file.");

        _stream.Position = _origin + ((long)sector + 1) * _sectorSize;
        _stream.ReadExactly(buffer);
    }

    private static int StreamSize(ReadOnlySpan<byte> entry, ReadOnlySpan<byte> header)
    {
        var field = entry[120..];
        ulong size = U16(header[26..]) == 3 ? U32(field) : BinaryPrimitives.ReadUInt64LittleEndian(field);
        if (size > MaxWorkbookBytes)
            throw Invalid("Compound file stream is too large.");

        return (int)size;
    }

    private static ushort U16(ReadOnlySpan<byte> data) => BinaryPrimitives.ReadUInt16LittleEndian(data);

    private static uint U32(ReadOnlySpan<byte> data) => BinaryPrimitives.ReadUInt32LittleEndian(data);

    private static InvalidDataException Invalid(string message) => new(message);
}
