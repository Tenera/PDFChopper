using System;
using System.IO;
using FluentAssertions;
using PdfChopper.Services;
using Xunit;

namespace PdfChopper.Tests.Services;

public class PdfServiceGetPageCountTests : IDisposable
{
    private readonly string _tempDir = TestHelper.CreateTempDir();
    private readonly PdfService _sut = new();

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public void GetPageCount_ReturnsCorrectCount()
    {
        var path = TestHelper.CreateTestPdf(_tempDir, 7);

        _sut.GetPageCount(path).Should().Be(7);
    }

    [Fact]
    public void GetPageCount_InvalidPath_Throws()
    {
        var act = () => _sut.GetPageCount(Path.Combine(_tempDir, "nonexistent.pdf"));

        act.Should().Throw<Exception>();
    }
}
