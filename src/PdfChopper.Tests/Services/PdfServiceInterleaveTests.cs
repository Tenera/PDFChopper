using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using FluentAssertions;
using PdfChopper.Models;
using PdfChopper.Services;
using Xunit;

namespace PdfChopper.Tests.Services;

public class PdfServiceInterleaveTests : IDisposable
{
    private readonly string _tempDir = TestHelper.CreateTempDir();

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public async Task InterleaveAsync_TwoEqualFiles_ProducesCorrectPageCount()
    {
        var file1 = TestHelper.CreateTestPdf(_tempDir, 2);
        var file2 = TestHelper.CreateTestPdf(_tempDir, 2);
        var output = Path.Combine(_tempDir, "interleaved.pdf");

        var files = new List<PdfFile> { new(file1), new(file2) };

        await PdfService.InterleaveAsync(files, output);

        TestHelper.GetPageCount(output).Should().Be(4);
    }

    [Fact]
    public async Task InterleaveAsync_UnequalPageCounts_IncludesAllPages()
    {
        var file1 = TestHelper.CreateTestPdf(_tempDir, 3);
        var file2 = TestHelper.CreateTestPdf(_tempDir, 1);
        var output = Path.Combine(_tempDir, "interleaved.pdf");

        var files = new List<PdfFile> { new(file1), new(file2) };

        await PdfService.InterleaveAsync(files, output);

        // Round-robin: A1, B1, A2, A3 = 4 pages total
        TestHelper.GetPageCount(output).Should().Be(4);
    }
}
