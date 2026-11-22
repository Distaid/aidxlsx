using System.Diagnostics;
using System.IO;
using System.Linq;
using System;
using Aid.Xlsx.CodeGen.Model.Elements;
using Aid.Xlsx.CodeGen.Model.TopLevel;
using Aid.Xlsx.CodeGen.Model;
using Aid.Xlsx.CodeGen.XsdParser;

namespace Aid.Xlsx.CodeGen;

public static class Program
{
    public static void Main(string[] args)
    {
        if (args.Length != 3)
        {
            Console.Error.WriteLine("Usage:");
            Console.Error.WriteLine($"    {Process.GetCurrentProcess().ProcessName}.exe command name-of-ooxml.xsd output.cs");
            Console.Error.WriteLine();
            return;
        }

        var command = args[0];
        var schemaPath = args[1];
        var target = args[2];
        using var fileStream = File.OpenRead(schemaPath);
        using var reader = new XmlTreeReader(fileStream, new XsdEnumMapper(), true);
        var parser = new XsdSchemaParser();

        var schema = parser.ParseSchema(reader);
        switch (command)
        {
            case "styles":
                GenerateStylesReader(schema, target);
                break;

            case "cache-records":
                GenerateCacheRecords(schema, target);
                break;

            case "partial-rst":
                GeneratePartialRstReader(schema, target);
                break;

            case "partial-sheet-data":
                GeneratePartialSheetDataReader(schema, target);
                break;

            default:
                Console.WriteLine($"Unknown command '{command}'");
                break;
        }
    }

    private static void GenerateStylesReader(Schema schema, string target)
    {
        var typeMap = new SchemeTypeMap()
            .AddPrimitiveTypes()
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_NumFmtId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt",
                OptionalTemplate = "_reader.GetOptionalUInt"
            })
            .AddStFontId()
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_FillId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt",
                OptionalTemplate = "_reader.GetOptionalUInt"
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_BorderId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt",
                OptionalTemplate = "_reader.GetOptionalUInt"
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_CellStyleXfId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt",
                OptionalTemplate = "_reader.GetOptionalUInt"
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_TextRotation",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt",
                OptionalTemplate = "_reader.GetOptionalUInt"
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_DxfId",
                CsTypeName = "uint",
                RequiredTemplate = "_reader.GetUInt",
                OptionalTemplate = "_reader.GetOptionalUInt"
            })
            .AddSimpleTypeEnum("ST_PatternType", "XlsxFillPatternValues")
            .AddSimpleTypeEnum("ST_GradientType", "XlsxGradientType", "linear", "XlsxGradientType.Linear")
            .AddSimpleTypeEnum("ST_BorderStyle", "XlsxBorderStyleValues", "none", "XlsxBorderStyleValues.None")
            .AddSimpleTypeEnum("ST_HorizontalAlignment", "XlsxAlignmentHorizontalValues")
            .AddSimpleTypeEnum("ST_VerticalAlignment", "XlsxAlignmentVerticalValues", "bottom", "XlsxAlignmentVerticalValues.Bottom")
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_TableStyleType",
                CsTypeName = "(XlsxTableStyleRegionValues?, XlsxPivotStyleRegionValues?)",
                RequiredTemplate = "_reader.GetStringMappedValue(\"{0}\", TableStyleTypeMap)"
            })
            .AddComplexTypeMapping("CT_Color", "XlsxColor")
            .AddComplexTypeMapping("CT_GradientStop", "(FractionOfOne Value, XlsxColor Color)")
            .AddComplexTypeMapping("CT_Font", "XlsxDifferentialFontValue")
            .AddComplexTypeMapping("CT_Fill", "XlsxFillFormatValue")
            .AddComplexTypeMapping("CT_Border", "XlsxDifferentialBorderValue")
            .AddComplexTypeMapping("CT_BorderPr", "XlsxBorderLine")
            .AddComplexTypeMapping("CT_PatternFill", "XlsxFillFormatValue")
            .AddComplexTypeMapping("CT_GradientFill", "XlsxFillFormatValue")
            .AddComplexTypeMapping("CT_NumFmt", "(int NumFmtId, XlsxNumberFormat Format)")
            .AddComplexTypeMapping("CT_CellAlignment", "XlsxDifferentialAlignmentValue")
            .AddComplexTypeMapping("CT_CellProtection", "XlsxDifferentialProtectionValue")
            .AddComplexTypeMapping("CT_Xf", "(XlsxCellFormatValue Format, int? CellStyleXfId)")
            .AddComplexTypeMapping("CT_CellXfs", "List<(XlsxCellFormatValue Format, int? CellStyleXfId)>")
            .AddComplexTypeMapping("CT_CellStyle", "(int CellStyleXfId, XlsxCellStyleValue Style)")
            .AddComplexTypeMapping("CT_CellStyles", "Dictionary<int, XlsxCellStyleValue>")
            .AddComplexTypeMapping("CT_RgbColor", "uint")
            ;

        var stylesReaderGenerator = new ParserGenerator(schema, typeMap, "StylesReader")
            .AddUsing("System.Collections.Generic")
            .AddUsing("Aid.Xlsx.Coordinates")
            .AddUsing("Aid.Xlsx.Opc")
            .AddUsing("Aid.Xlsx.Styles")
            .AddUsing("Aid.Xlsx.Styles.Colors")
            .AddUsing("Aid.Xlsx.Styles.Formatting")
            //.AddParseMethod("CT_Stylesheet")
            .AddParseMethod("CT_NumFmts")
            .AddParseMethod("CT_NumFmt")
            .AddParseMethod("CT_Fonts")
            // AddParseMethod("CT_Font")
            .AddParseMethod("CT_Fills")
            .AddParseMethod("CT_Fill")
            .AddParseMethod("CT_PatternFill")
            .AddParseMethod("CT_GradientFill")
            .AddParseMethod("CT_GradientStop")
            .AddParseMethod("CT_Borders")
            .AddParseMethod("CT_Border")
            .AddParseMethod("CT_BorderPr")
            .AddParseMethod("CT_CellStyleXfs")
            .AddParseMethod("CT_Xf")
            .AddParseMethod("CT_CellAlignment")
            .AddParseMethod("CT_CellProtection")
            .AddParseMethod("CT_CellXfs")
            .AddParseMethod("CT_CellStyles")
            .AddParseMethod("CT_CellStyle")
            .AddParseMethod("CT_Dxfs")
            .AddParseMethod("CT_Dxf")
            .AddParseMethod("CT_TableStyles")
            .AddParseMethod("CT_TableStyle")
            .AddParseMethod("CT_TableStyleElement")
            .AddParseMethod("CT_Colors")
            .AddParseMethod("CT_IndexedColors")
            .AddParseMethod("CT_MRUColors")
            .AddParseMethod("CT_RgbColor")
            ;

        var stylesReaderSource = stylesReaderGenerator.Generate();
        File.WriteAllText(target, stylesReaderSource);
        Console.WriteLine(stylesReaderSource);
    }

    private static void GenerateCacheRecords(Schema schema, string target)
    {
        var typeMap = new SchemeTypeMap()
            .AddPrimitiveTypes();

        var cacheRecordsGenerator = new ParserGenerator(schema, typeMap, "PivotCacheRecordsReader")
            .WithNamespace("Aid.Xlsx.Internal.IO")
            .AddUsing("System.Collections.Generic")
            .AddUsing("Aid.Xlsx.Opc")

            // CT_PivotCacheRecords - hand-coded
            .AddParseMethod("CT_Record")
            .AddParseMethod("CT_Missing")
            .AddParseMethod("CT_Number")
            .AddParseMethod("CT_Boolean")
            .AddParseMethod("CT_Error")
            .AddParseMethod("CT_String")
            .AddParseMethod("CT_DateTime")
            .AddParseMethod("CT_Index")
            .AddParseMethod("CT_X")
            .AddParseMethod("CT_Tuples")
            .AddParseMethod("CT_Tuple")
            ;

        var cacheRecordsSource = cacheRecordsGenerator.Generate();
        File.WriteAllText(target, cacheRecordsSource);
        Console.WriteLine(cacheRecordsSource);
    }

    /// <summary>
    /// Generate a reader to read <c>CT_Rst</c>. It is a partial reader used in other readers.
    /// </summary>
    private static void GeneratePartialRstReader(Schema schema, string target)
    {
        NormalizeChoicesZeroToOne(schema, "CT_RPrElt");

        var typeMap = new SchemeTypeMap()
            .AddPrimitiveTypes()
            .AddStFontId()
            .AddSimpleTypeEnum("ST_PhoneticType", "XlsxPhoneticType", "fullwidthKatakana", "XlsxPhoneticType.FullWidthKatakana")
            .AddSimpleTypeEnum("ST_PhoneticAlignment", "XlsxPhoneticAlignment", "left", "XlsxPhoneticAlignment.Left")
            .AddSimpleTypeEnum("s:ST_VerticalAlignRun", "XlsxFontVerticalTextAlignmentValues")
            .AddSimpleTypeEnum("ST_FontScheme", "XlsxFontScheme")
            .AddSimpleTypeEnum("ST_UnderlineValues", "XlsxFontUnderlineValues", "single", "XlsxFontUnderlineValues.Single")

            .AddComplexTypeMapping("CT_Color", "XlsxColor")
            .AddComplexTypeMapping("CT_BooleanProperty", "bool")
            .AddComplexTypeMapping("CT_IntProperty", "int")
            .AddComplexTypeMapping("CT_FontSize", "XlsxFontSize")
            .AddComplexTypeMapping("CT_UnderlineProperty", "XlsxFontUnderlineValues")
            .AddComplexTypeMapping("CT_VerticalAlignFontProperty", "XlsxFontVerticalTextAlignmentValues")
            .AddComplexTypeMapping("CT_FontScheme", "XlsxFontScheme")
            .AddComplexTypeMapping("CT_FontName", "XlsxFontName")

            .AddComplexTypeMapping("s:ST_Xstring", "string")
            .AddComplexTypeMapping("CT_RPrElt", "XlsxDifferentialFontValue", "Unit")
            .AddComplexTypeMapping("CT_RElt", "(string Text, XlsxDifferentialFontValue Font)")
            .AddComplexTypeMapping("CT_PhoneticRun", "PhoneticRunDto")
            .AddComplexTypeMapping("CT_PhoneticPr", "PhoneticProperties")
            .AddComplexTypeMapping("CT_Rst", "OneOf<string, XlsxImmutableRichText>")

            .AddParseCall("s:ST_Xstring", "_reader.ParseXString")
            ;

        var rstGenerator = new ParserGenerator(schema, typeMap, "RstReader")
            .WithNamespace("Aid.Xlsx.Internal.IO")
            .AddUsing("System.Collections.Generic")
            .AddUsing("Aid.Xlsx.CalcEngine")
            .AddUsing("Aid.Xlsx.Opc")
            .AddUsing("Aid.Xlsx.RichText")
            .AddUsing("Aid.Xlsx.Styles")
            .AddUsing("Aid.Xlsx.Styles.Formatting")
            .AddUsing("PhoneticProperties = Aid.Xlsx.RichText.XlsxImmutableRichText.PhoneticProperties")
            .AddUsing("PhoneticRunDto = (string Text, int StartIndex, int EndIndex)")

            .AddParseMethod("CT_PhoneticRun") // 1816
            .AddParseMethod("CT_RElt") // 1823
            .AddParseMethod("CT_RPrElt") // 1829
            .AddParseMethod("CT_Rst") // 1849
            .AddParseMethod("CT_PhoneticPr") // 1857
            .AddParseMethod("CT_BooleanProperty") // 3751
            .AddParseMethod("CT_FontSize") // 3754
            .AddParseMethod("CT_IntProperty") // 3757
            .AddParseMethod("CT_FontName") // 3760
            .AddParseMethod("CT_VerticalAlignFontProperty") // 3763
            .AddParseMethod("CT_FontScheme") // 3766
            .AddParseMethod("CT_UnderlineProperty") // 3776
            ;

        var rstSource = rstGenerator.Generate();
        File.WriteAllText(target, rstSource);
        Console.WriteLine(rstSource);
    }

    /// <summary>
    /// Generate a partial reader to read <c>CT_SheetData</c>.
    /// </summary>
    private static void GeneratePartialSheetDataReader(Schema schema, string target)
    {
        // Add <row> ac:dyDescent attribute. It's one of few MCE attribute
        var row = schema.Entries.OfType<ComplexTypeSequence>().Single(x => x.Name == "CT_Row");
        row.Attributes.Add(new AttributeElement
        {
            Name = "dyDescent",
            NsName = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac",
            Type = "xsd:double",
            RefName = null,
        });

        // ST_CellSpans is a list of spans the cell uses in the row, but is unreliable and can be just omitted
        row.Attributes.RemoveAll(x => x.TryPickT1(out var attr, out _) && attr.Name == "spans");

        var typeMap = new SchemeTypeMap()
            .AddPrimitiveTypes()
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_CellRef",
                CsTypeName = "Point",
                OptionalTemplate = "_reader.GetOptionalPoint"
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_CellType",
                CsTypeName = "string",
                RequiredTemplate = "_reader.GetString",
                OptionalTemplate = "_reader.GetOptionalString",
                MapValue = x => x == "n" ? "\"n\"" : throw new NotSupportedException()
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                // Keep "" for default value of ST_CellFormulaType so normal type can be checked by a empty string check
                Name = "ST_CellFormulaType",
                CsTypeName = "string",
                OptionalTemplate = "_reader.GetOptionalString",
                MapValue = x => x == "normal" ? "string.Empty" : throw new NotSupportedException()
            })
            .AddSimpleType(new SimpleTypeMapping
            {
                Name = "ST_Ref",
                CsTypeName = "Area",
                OptionalTemplate = "_reader.GetOptionalArea"
            })

            .AddComplexTypeMapping("s:ST_Xstring", "string")
            .AddComplexTypeMapping("CT_Rst", "StringItem")

            .AddParseCall("s:ST_Xstring", "_reader.ParseXString")
            .AddParseCall("CT_Rst", "_rstReader.ParseCtRst")
            ;

        var sheetDataGenerator = new ParserGenerator(schema, typeMap, "SheetDataReader")
            .WithNamespace("Aid.Xlsx.Internal.IO")
            .AddUsing("System.Collections.Generic")
            .AddUsing("Aid.Xlsx.Coordinates")
            .AddUsing("Aid.Xlsx.Opc")
            .AddUsing("StringItem = Aid.Xlsx.CalcEngine.OneOf<string, Aid.Xlsx.RichText.XlsxImmutableRichText>")

            .AddParseMethod("CT_SheetData") // 2232
            .AddParseMethod("CT_Row") // 2274
            .AddParseMethod("CT_Cell") // 2292
            .AddParseMethod("CT_CellFormula") // 2772
            ;

        var sheetDataSource = sheetDataGenerator.Generate();
        File.WriteAllText(target, sheetDataSource);
        Console.WriteLine(sheetDataSource);
    }

    /// <summary>
    /// Normalize a choice that has multiple elements with 0..1 cardinality into the 1..1
    /// cardinality. Semantically, it's identical.
    /// </summary>
    private static void NormalizeChoicesZeroToOne(Schema schema, ParsletName name)
    {
        foreach (var choice in schema.Entries.OfType<ComplexTypeChoice>().Where(x => name == x.Name))
        {
            var choiceChildren = choice.Choice.Children;
            var exhibitsProblem = choiceChildren.All(x => x is ElementType
            {
                Occurrences: { ActualMin: 0, ActualMax: 1 }
            });
            if (!exhibitsProblem)
            {
                throw new InvalidOperationException($"Type {name} does not exhibit expected problem.");
            }

            var fixedChildren = choiceChildren.Cast<ElementType>().Select(x => new ElementType
            {
                Name = x.Name,
                Occurrences = new Occurrences(1, 1),
                TypeName = x.TypeName
            }).ToList();

            // Replace the original children with new ones
            choice.Choice.Children.Clear();
            choice.Choice.Children.AddRange(fixedChildren);
        }
    }

    public static SchemeTypeMap AddStFontId(this SchemeTypeMap schemeTypeMap)
    {
        return schemeTypeMap.AddSimpleType(new SimpleTypeMapping
        {
            Name = "ST_FontId",
            CsTypeName = "uint",
            RequiredTemplate = "_reader.GetUInt",
            OptionalTemplate = "_reader.GetOptionalUInt"
        });
    }
}
