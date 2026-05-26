using System;
using System.Collections.Generic;
using System.IO;
using FluentAssertions;
using PdfChopper.Models;
using PdfChopper.Services;
using Xunit;

namespace PdfChopper.Tests.Services;

public class PdfServiceSplitTests : IDisposable
{
    private readonly string _tempDir = TestHelper.CreateTempDir();
    private readonly PdfService _sut = new();

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public void Split_TwoExtracts_CreatesCorrectFiles()
    {
        var inputPath = TestHelper.CreateTestPdf(_tempDir, 4);
        var pageCount = TestHelper.GetPageCount(inputPath);

        var extract1 = new PdfFileExtract(pageCount, Path.Combine(_tempDir, "part1.pdf"));
        extract1.StartPage = 1;
        extract1.EndPage = 2;

        var extract2 = new PdfFileExtract(pageCount, Path.Combine(_tempDir, "part2.pdf"));
        extract2.StartPage = 3;
        extract2.EndPage = 4;

        _sut.Split(inputPath, new List<PdfFileExtract> { extract1, extract2 });

        TestHelper.GetPageCount(extract1.FilePath).Should().Be(2);
        TestHelper.GetPageCount(extract2.FilePath).Should().Be(2);
    }

    [Fact]
    public void Split_SinglePageExtract_Works()
    {
        var inputPath = TestHelper.CreateTestPdf(_tempDir, 3);
        var pageCount = TestHelper.GetPageCount(inputPath);

        var extract = new PdfFileExtract(pageCount, Path.Combine(_tempDir, "single.pdf"));
        extract.StartPage = 2;
        extract.EndPage = 2;

        _sut.Split(inputPath, new List<PdfFileExtract> { extract });

        TestHelper.GetPageCount(extract.FilePath).Should().Be(1);
    }
}
