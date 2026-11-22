using System;
using System.IO;
using Aid.Xlsx.CodeGen.Schema;
using Aid.Xlsx.CodeGen.SpreadsheetReaders;

namespace Aid.Xlsx.CodeGen;

public static class Program
{
    public static int Main(string[] args)
    {
        if (args is ["all"])
        {
            var root = OoxmlSchema.FindRepository();
            var schemaPath = OoxmlSchema.GetPath(root);
            var output = Path.Combine(root, "src", "Aid.Xlsx", "SpreadsheetML");
            ReaderRecipes.Generate("styles", schemaPath, Path.Combine(output, "Styles", "StylesReader.g.cs"));
            ReaderRecipes.Generate("cache-records", schemaPath, Path.Combine(output, "Pivots", "PivotCacheRecordsReader.g.cs"));
            ReaderRecipes.Generate("partial-rst", schemaPath, Path.Combine(output, "Strings", "RstReader.g.cs"));
            ReaderRecipes.Generate("partial-sheet-data", schemaPath, Path.Combine(output, "Worksheets", "SheetDataReader.g.cs"));
            return 0;
        }

        if (args is [var command, var input, var target])
            return ReaderRecipes.Generate(command, input, target);

        Console.Error.WriteLine("Usage: Aid.Xlsx.CodeGen all | <command> <schema.xsd> <output.cs>");
        return 1;
    }
}
