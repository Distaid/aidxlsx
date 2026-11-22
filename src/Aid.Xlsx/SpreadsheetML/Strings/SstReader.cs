using StringItem = Aid.Xlsx.Calculation.OneOf<string, Aid.Xlsx.Model.RichText.XlsxImmutableRichText>;
using Aid.Xlsx.Calculation;
using Aid.Xlsx.Model.Styles;
using Aid.Xlsx.Opc.Xml;
using Aid.Xlsx.SpreadsheetML.Xml;

namespace Aid.Xlsx.SpreadsheetML.Strings;

/// <summary>
/// A reader for shared string table part.
/// </summary>
internal class SstReader
{
    private readonly string _ns = OpenXmlConst.Main2006SsNs;
    private readonly XmlTreeReader _reader;
    private readonly RstReader _rstReader;

    internal SstReader(XmlTreeReader reader, XlsxWorkbookStyles styles)
    {
        _reader = reader;
        _rstReader = new RstReader(reader, styles);
    }

    internal List<StringItem> ParseSst()
    {
        if (!_reader.TryOpen("sst", _ns))
        {
            throw PartStructureException.ExpectedElementNotFound("sst", _reader);
        }

        // Do not rely on the supplied values, could allocate unlimited memory
        _ = _reader.GetOptionalUInt("count");
        _ = _reader.GetOptionalUInt("uniqueCount");

        var sst = new List<StringItem>();
        while (_rstReader.ParseCtRst("si", _ns) is { IsSuccess: true } si)
        {
            sst.Add(si.Value);
        }

        // SST has no extensions, skip
        if (_reader.TryOpen("extLst", _ns))
            _reader.Skip("extLst");

        _reader.Close("sst", _ns);

        return sst;
    }
}
