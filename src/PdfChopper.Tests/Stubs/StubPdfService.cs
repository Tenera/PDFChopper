using System.Collections.Generic;
using PdfChopper.Models;
using PdfChopper.Services;

namespace PdfChopper.Tests.Stubs;

public class StubPdfService : IPdfService
{
    public int PageCount { get; set; } = 5;

    public int GetPageCount(string filePath) => PageCount;
    public void Merge(IReadOnlyList<PdfFile> files, string outputPath) { }
    public void Split(string inputPath, IReadOnlyList<PdfFileExtract> extracts) { }
    public void Rotate(string inputPath, IReadOnlyList<PdfFileRotation> parts, string outputPath) { }
    public void Interleave(IReadOnlyList<PdfFile> files, string outputPath) { }
    public void Reorder(string inputPath, string outputPath) { }
}
