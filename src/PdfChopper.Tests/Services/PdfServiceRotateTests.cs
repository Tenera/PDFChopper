using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using FluentAssertions;
using PdfChopper.Models;
using PdfChopper.Services;
using PdfSharp.Pdf.IO;
using Xunit;

namespace PdfChopper.Tests.Services;

public class PdfServiceRotateTests : IDisposable
{
    private readonly string _tempDir = TestHelper.CreateTempDir();

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public async Task RotateAsync_RotatesSpecifiedPages()
    {
        var inputPath = TestHelper.CreateTestPdf(_tempDir, 3);
        var outputPath = Path.Combine(_tempDir, "rotated.pdf");
        var parent = new PdfFile(inputPath);

        var part = new PdfFileRotation(parent);
        part.StartPage = 1;
        part.EndPage = 2;
        part.Rotate = 1; // 90 degrees

        await PdfService.RotateAsync(inputPath, new List<PdfFileRotation> { part }, outputPath);

        using var result = PdfReader.Open(outputPath, PdfDocumentOpenMode.Modify);
        result.Pages[0].Rotate.Should().Be(90);
        result.Pages[1].Rotate.Should().Be(90);
        result.Pages[2].Rotate.Should().Be(0);
        result.Close();
    }

    [Fact]
    public async Task RotateAsync_ZeroRotation_LeavesUnchanged()
    {
        var inputPath = TestHelper.CreateTestPdf(_tempDir, 2);
        var outputPath = Path.Combine(_tempDir, "rotated.pdf");
        var parent = new PdfFile(inputPath);

        var part = new PdfFileRotation(parent);
        part.Rotate = 0; // no rotation (0 mod 4 = 0)

        await PdfService.RotateAsync(inputPath, new List<PdfFileRotation> { part }, outputPath);

        using var result = PdfReader.Open(outputPath, PdfDocumentOpenMode.Modify);
        result.Pages[0].Rotate.Should().Be(0);
        result.Pages[1].Rotate.Should().Be(0);
        result.Close();
    }

    [Fact]
    public async Task RotateAsync_RotationWraps_Mod4()
    {
        var inputPath = TestHelper.CreateTestPdf(_tempDir, 1);
        var outputPath = Path.Combine(_tempDir, "rotated.pdf");
        var parent = new PdfFile(inputPath);

        var part = new PdfFileRotation(parent);
        part.Rotate = 5; // 5 mod 4 = 1, so 90 degrees

        await PdfService.RotateAsync(inputPath, new List<PdfFileRotation> { part }, outputPath);

        using var result = PdfReader.Open(outputPath, PdfDocumentOpenMode.Modify);
        result.Pages[0].Rotate.Should().Be(90);
        result.Close();
    }
}
