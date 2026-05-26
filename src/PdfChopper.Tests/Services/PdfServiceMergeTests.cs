using System;
using System.Collections.Generic;
using System.IO;
using FluentAssertions;
using PdfChopper.Models;
using PdfChopper.Services;
using Xunit;

namespace PdfChopper.Tests.Services;

public class PdfServiceMergeTests : IDisposable
{
    private readonly string _tempDir = TestHelper.CreateTempDir();
    private readonly PdfService _sut = new();

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public void Merge_TwoFiles_ProducesCorrectPageCount()
    {
        var file1 = TestHelper.CreateTestPdf(_tempDir, 3);
        var file2 = TestHelper.CreateTestPdf(_tempDir, 2);
        var output = Path.Combine(_tempDir, "merged.pdf");

        var files = new List<PdfFile> { new(file1, 3), new(file2, 2) };

        _sut.Merge(files, output);

        TestHelper.GetPageCount(output).Should().Be(5);
    }

    [Fact]
    public void Merge_WithPageRanges_MergesOnlySelectedPages()
    {
        var file1 = TestHelper.CreateTestPdf(_tempDir, 4);
        var output = Path.Combine(_tempDir, "merged.pdf");

        var pdfFile = new PdfFile(file1, 4);
        pdfFile.StartPage = 2;
        pdfFile.EndPage = 3;

        _sut.Merge(new List<PdfFile> { pdfFile }, output);

        TestHelper.GetPageCount(output).Should().Be(2);
    }

    [Fact]
    public void Merge_SingleFile_CopiesAllPages()
    {
        var file = TestHelper.CreateTestPdf(_tempDir, 5);
        var output = Path.Combine(_tempDir, "merged.pdf");

        _sut.Merge(new List<PdfFile> { new(file, 5) }, output);

        TestHelper.GetPageCount(output).Should().Be(5);
    }
}
