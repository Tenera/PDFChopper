using System;
using System.Collections.Generic;
using PdfChopper.Models;
using PdfChopper.Services;

namespace PdfChopper.Tests.Stubs;

public class StubPdfService : IPdfService
{
    public int PageCount { get; set; } = 5;
    public Exception? ExceptionToThrow { get; set; }

    private void ThrowIfConfigured()
    {
        if (ExceptionToThrow is not null)
            throw ExceptionToThrow;
    }

    public int GetPageCount(string filePath) { ThrowIfConfigured(); return PageCount; }
    public void Merge(IReadOnlyList<PdfFile> files, string outputPath) => ThrowIfConfigured();
    public void Split(string inputPath, IReadOnlyList<PdfFileExtract> extracts) => ThrowIfConfigured();
    public void Rotate(string inputPath, IReadOnlyList<PdfFileRotation> parts, string outputPath) => ThrowIfConfigured();
    public void Interleave(IReadOnlyList<PdfFile> files, string outputPath) => ThrowIfConfigured();
    public void Reorder(string inputPath, string outputPath) => ThrowIfConfigured();
}
