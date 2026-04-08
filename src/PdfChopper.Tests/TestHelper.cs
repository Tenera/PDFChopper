using System;
using System.IO;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace PdfChopper.Tests;

public static class TestHelper
{
    public static string CreateTestPdf(string directory, int pageCount, string? fileName = null)
    {
        fileName ??= $"{Guid.NewGuid()}.pdf";
        var path = Path.Combine(directory, fileName);
        using var doc = new PdfDocument();
        for (var i = 0; i < pageCount; i++)
        {
            doc.AddPage(new PdfPage());
        }
        doc.Save(path);
        return path;
    }

    public static int GetPageCount(string path)
    {
        using var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import);
        return doc.PageCount;
    }

    public static string CreateTempDir()
    {
        var dir = Path.Combine(Path.GetTempPath(), "PdfChopperTests", Guid.NewGuid().ToString());
        Directory.CreateDirectory(dir);
        return dir;
    }
}
