using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using FluentAssertions;
using PdfChopper.Models;
using PdfChopper.Services;
using Xunit;

namespace PdfChopper.Tests.Services;

public class PdfServiceMergeTests : IDisposable
{
    private readonly string _tempDir = TestHelper.CreateTempDir();

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public async Task MergeAsync_TwoFiles_ProducesCorrectPageCount()
    {
        var file1 = TestHelper.CreateTestPdf(_tempDir, 3);
        var file2 = TestHelper.CreateTestPdf(_tempDir, 2);
        var output = Path.Combine(_tempDir, "merged.pdf");

        var files = new List<PdfFile> { new(file1), new(file2) };

        await PdfService.MergeAsync(files, output);

        TestHelper.GetPageCount(output).Should().Be(5);
    }

    [Fact]
    public async Task MergeAsync_WithPageRanges_MergesOnlySelectedPages()
    {
        var file1 = TestHelper.CreateTestPdf(_tempDir, 4);
        var output = Path.Combine(_tempDir, "merged.pdf");

        var pdfFile = new PdfFile(file1);
        pdfFile.StartPage = 2;
        pdfFile.EndPage = 3;

        await PdfService.MergeAsync(new List<PdfFile> { pdfFile }, output);

        TestHelper.GetPageCount(output).Should().Be(2);
    }

    [Fact]
    public async Task MergeAsync_SingleFile_CopiesAllPages()
    {
        var file = TestHelper.CreateTestPdf(_tempDir, 5);
        var output = Path.Combine(_tempDir, "merged.pdf");

        await PdfService.MergeAsync(new List<PdfFile> { new(file) }, output);

        TestHelper.GetPageCount(output).Should().Be(5);
    }
}
