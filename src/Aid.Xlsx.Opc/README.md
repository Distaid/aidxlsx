# Aid.Xlsx.Opc

This project implements Open Packaging Conventions (OPC) and XML primitives used by the workbook
reader, writer, and schema generator. The main NuGet package includes the OPC assembly.

| Directory | Responsibility |
| --- | --- |
| [Packaging](Packaging/) | ZIP entries, part names and content types, package and part relationships, and limits applied when opening a package |
| [Xml](Xml/) | Forward-only XML traversal, attribute conversion, and SpreadsheetML string escape decoding |
| [Xml/Compatibility](Xml/Compatibility/) | Markup Compatibility and Extensibility (MCE) processing, including alternate content |

[`OpcPackage`](Packaging/OpcPackage.cs) is the entry point for package I/O. It exposes parts and
relationships and reads payloads on demand. [`XmlTreeReader`](Xml/XmlTreeReader.cs) is the entry
point for forward-only XML parsing. Workbook-specific interpretation lives in the SpreadsheetML
codecs in the main project.
