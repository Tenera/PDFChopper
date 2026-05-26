using System;
using System.Collections.Generic;
using System.IO;
using FluentAssertions;
using PdfChopper.Models;
using PdfChopper.Services;
using PdfSharp.Pdf.IO;
using Xunit;

namespace PdfChopper.Tests.Services;

public class PdfServiceRotateTests : IDisposable
{
    private readonly string _tempDir = TestHelper.CreateTempDir();
    private readonly PdfService _sut = new();

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    [Fact]
    public void Rotate_RotatesSpecifiedPages()
    {
        var inputPath = TestHelper.CreateTestPdf(_tempDir, 3);
        var outputPath = Path.Combine(_tempDir, "rotated.pdf");
        var pageCount = TestHelper.GetPageCount(inputPath);

        var part = new PdfFileRotation(pageCount);
        part.StartPage = 1;
        part.EndPage = 2;
        part.Rotate = 1;

        _sut.Rotate(inputPath, new List<PdfFileRotation> { part }, outputPath);

        using var result = PdfReader.Open(outputPath, PdfDocumentOpenMode.Modify);
        result.Pages[0].Rotate.Should().Be(90);
        result.Pages[1].Rotate.Should().Be(90);
        result.Pages[2].Rotate.Should().Be(0);
    }

    [Fact]
    public void Rotate_ZeroRotation_LeavesUnchanged()
    {
        var inputPath = TestHelper.CreateTestPdf(_tempDir, 2);
        var outputPath = Path.Combine(_tempDir, "rotated.pdf");
        var pageCount = TestHelper.GetPageCount(inputPath);

        var part = new PdfFileRotation(pageCount);
        part.Rotate = 0;

        _sut.Rotate(inputPath, new List<PdfFileRotation> { part }, outputPath);

        using var result = PdfReader.Open(outputPath, PdfDocumentOpenMode.Modify);
        result.Pages[0].Rotate.Should().Be(0);
        result.Pages[1].Rotate.Should().Be(0);
    }

    [Fact]
    public void Rotate_RotationWraps_Mod4()
    {
        var inputPath = TestHelper.CreateTestPdf(_tempDir, 1);
        var outputPath = Path.Combine(_tempDir, "rotated.pdf");
        var pageCount = TestHelper.GetPageCount(inputPath);

        var part = new PdfFileRotation(pageCount);
        part.Rotate = 5;

        _sut.Rotate(inputPath, new List<PdfFileRotation> { part }, outputPath);

        using var result = PdfReader.Open(outputPath, PdfDocumentOpenMode.Modify);
        result.Pages[0].Rotate.Should().Be(90);
    }
}
