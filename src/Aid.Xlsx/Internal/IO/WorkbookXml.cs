using System.Xml.Linq;
using System.Xml;
using Aid.Xlsx.Protection;
using Aid.Xlsx.Utils;
using Aid.Xlsx.Workbooks;
using Aid.Xlsx.Worksheets;

namespace Aid.Xlsx.Internal.IO;

internal static class WorkbookXml
{
    internal static XElement ReadRoot(Stream stream)
    {
        using var reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        var root = XElement.Load(reader);
        if (root.Name.LocalName != "workbook" || root.Name.NamespaceName is not (OpenXmlConst.Main2006SsNs or OpenXmlConst.StrictSpreadsheetNs))
            throw new InvalidDataException("The workbook part does not contain a SpreadsheetML workbook.");

        return root;
    }

    internal static void LoadSettings(XElement root, XlsxWorkbook workbook)
    {
        var ns = root.Name.Namespace;
        workbook.Use1904DateSystem = (bool?)root.Element(ns + "workbookPr")?.Attribute("date1904") ?? false;

        if (root.Element(ns + "fileSharing") is { } sharing)
        {
            workbook.FileSharing.ReadOnlyRecommended = (bool?)sharing.Attribute("readOnlyRecommended") ?? false;
            workbook.FileSharing.UserName = (string?)sharing.Attribute("userName");
        }

        if (root.Element(ns + "workbookProtection") is { } protection)
            LoadProtection(protection, workbook);

        if (root.Element(ns + "calcPr") is not { } calculation)
            return;

        workbook.CalculateMode = (string?)calculation.Attribute("calcMode") switch
        {
            "auto" => XlsxCalculateMode.Auto,
            "autoNoTable" => XlsxCalculateMode.AutoNoTable,
            "manual" => XlsxCalculateMode.Manual,
            _ => workbook.CalculateMode
        };
        workbook.ReferenceStyle = (string?)calculation.Attribute("refMode") switch
        {
            "A1" => XlsxReferenceStyle.A1,
            "R1C1" => XlsxReferenceStyle.R1C1,
            _ => workbook.ReferenceStyle
        };
        workbook.CalculationOnSave = (bool?)calculation.Attribute("calcOnSave") ?? workbook.CalculationOnSave;
        workbook.ForceFullCalculation = (bool?)calculation.Attribute("forceFullCalc") ?? workbook.ForceFullCalculation;
        workbook.FullCalculationOnLoad = (bool?)calculation.Attribute("fullCalcOnLoad") ?? workbook.FullCalculationOnLoad;
        workbook.FullPrecision = (bool?)calculation.Attribute("fullPrecision") ?? workbook.FullPrecision;
    }

    internal static Sheet ReadSheet(XElement sheet)
    {
        var relation = (string?)sheet.Attribute(XName.Get("id", OpenXmlConst.RelationshipsNs))
            ?? (string?)sheet.Attribute(XName.Get("id", OpenXmlConst.StrictRelationshipsNs));
        var visibility = (string?)sheet.Attribute("state") switch
        {
            "hidden" => XlsxWorksheetVisibility.Hidden,
            "veryHidden" => XlsxWorksheetVisibility.VeryHidden,
            _ => XlsxWorksheetVisibility.Visible
        };
        return new Sheet((string?)sheet.Attribute("name") ?? throw new InvalidDataException("A workbook sheet has no name."),
            (uint?)sheet.Attribute("sheetId") ?? throw new InvalidDataException("A workbook sheet has no sheet ID."),
            relation, visibility);
    }

    private static void LoadProtection(XElement element, XlsxWorkbook workbook)
    {
        var protection = workbook.Protection;
        protection.IsProtected = true;
        var algorithm = (string?)element.Attribute("workbookAlgorithmName");
        if (string.IsNullOrEmpty(algorithm))
        {
            protection.PasswordHash = (string?)element.Attribute("workbookPassword") ?? string.Empty;
            protection.Base64EncodedSalt = string.Empty;
        }
        else if (XlsxProtectionAlgorithm.TryParseDescription(algorithm, out var parsedAlgorithm))
        {
            protection.Algorithm = parsedAlgorithm;
            protection.PasswordHash = (string?)element.Attribute("workbookHashValue") ?? string.Empty;
            protection.SpinCount = (uint?)element.Attribute("workbookSpinCount") ?? 0;
            protection.Base64EncodedSalt = (string?)element.Attribute("workbookSaltValue") ?? string.Empty;
        }

        protection.AllowElement(XlsxWorkbookProtectionElements.Structure, !((bool?)element.Attribute("lockStructure") ?? false));
        protection.AllowElement(XlsxWorkbookProtectionElements.Windows, !((bool?)element.Attribute("lockWindows") ?? false));
    }

    internal readonly record struct Sheet(string Name, uint SheetId, string? Id, XlsxWorksheetVisibility Visibility);
}
