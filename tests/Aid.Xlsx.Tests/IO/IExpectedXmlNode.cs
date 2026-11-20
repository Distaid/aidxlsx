#nullable enable
using Aid.Xlsx.Opc;

namespace Aid.Xlsx.Tests.IO;

internal interface IExpectedXmlNode
{
    void AssertMatches(IXmlReader reader);
}
