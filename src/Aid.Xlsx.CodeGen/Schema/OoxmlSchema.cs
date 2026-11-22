using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Security.Cryptography;

namespace Aid.Xlsx.CodeGen;

internal static class OoxmlSchema
{
    private const string ArchiveUrl = "https://ecma-international.org/wp-content/uploads/ECMA-376-4_5th_edition_december_2016.zip";
    private const string SchemaFile = "sml.xsd";

    internal static string FindRepository()
    {
        for (var directory = new DirectoryInfo(Directory.GetCurrentDirectory()); directory is not null; directory = directory.Parent)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Aid.Xlsx.slnx")))
                return directory.FullName;
        }

        throw new DirectoryNotFoundException("Run code generation from the Aid.Xlsx repository.");
    }

    internal static string GetPath(string repository)
    {
        var checksum = File.ReadAllText(Path.Combine(repository, "schemas", "ooxml", "SHA256SUMS")).Split(' ', 2)[0];
        var cache = Path.Combine(repository, "artifacts", "codegen");
        var schemaPath = Path.Combine(cache, SchemaFile);
        if (File.Exists(schemaPath))
        {
            VerifyHash(File.ReadAllBytes(schemaPath), checksum);
            return schemaPath;
        }

        using var client = new HttpClient();
        using var request = new HttpRequestMessage(HttpMethod.Get, ArchiveUrl);
        using var response = client.Send(request);
        response.EnsureSuccessStatusCode();
        using var source = response.Content.ReadAsStream();
        using var archive = new ZipArchive(source, ZipArchiveMode.Read);
        using var schemaArchiveStream = OpenEntry(archive, "OfficeOpenXML-XMLSchema-Transitional.zip");
        using var schemaArchive = new ZipArchive(schemaArchiveStream, ZipArchiveMode.Read);
        using var schema = OpenEntry(schemaArchive, SchemaFile);
        using var buffer = new MemoryStream();
        schema.CopyTo(buffer);
        var bytes = buffer.ToArray();
        VerifyHash(bytes, checksum);

        Directory.CreateDirectory(cache);
        var temporaryPath = Path.Combine(cache, Path.GetRandomFileName());
        try
        {
            File.WriteAllBytes(temporaryPath, bytes);
            File.Move(temporaryPath, schemaPath, overwrite: true);
        }
        finally
        {
            File.Delete(temporaryPath);
        }

        return schemaPath;
    }

    private static Stream OpenEntry(ZipArchive archive, string name) =>
        archive.GetEntry(name)?.Open() ?? throw new InvalidDataException($"The ECMA archive is missing '{name}'.");

    private static void VerifyHash(byte[] bytes, string checksum)
    {
        if (!Convert.ToHexString(SHA256.HashData(bytes)).Equals(checksum, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException($"The SHA-256 checksum of '{SchemaFile}' differs from schemas/ooxml/SHA256SUMS.");
    }
}
