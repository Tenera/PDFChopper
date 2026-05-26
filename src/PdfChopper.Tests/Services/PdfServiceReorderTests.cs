using System;
using System.IO;
using FluentAssertions;
using PdfChopper.Services;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Xunit;

namespace PdfChopper.Tests.Services;

public class PdfServiceReorderTests : IDisposable
{
    private readonly string _tempDir = TestHelper.CreateTempDir();
    private readonly PdfService _sut = new();

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    private string CreateIdentifiablePdf(int pageCount)
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid()}.pdf");
        using var doc = new PdfDocument();
        for (var i = 1; i <= pageCount; i++)
        {
            var page = new PdfPage { Width = XUnit.FromPoint(100 + i) };
            doc.AddPage(page);
        }
        doc.Save(path);
        return path;
    }

    private static double GetPageWidth(string path, int pageIndex)
    {
        using var doc = PdfReader.Open(path, PdfDocumentOpenMode.Import);
        return doc.Pages[pageIndex].Width.Point;
    }

    [Fact]
    public void Reorder_SinglePage_ReturnsUnchanged()
    {
        var input = CreateIdentifiablePdf(1);
        var output = Path.Combine(_tempDir, "reordered.pdf");

        _sut.Reorder(input, output);

        TestHelper.GetPageCount(output).Should().Be(1);
        GetPageWidth(output, 0).Should().Be(101);
    }

    [Fact]
    public void Reorder_TwoPages_InterleavesCorrectly()
    {
        var input = CreateIdentifiablePdf(2);
        var output = Path.Combine(_tempDir, "reordered.pdf");

        _sut.Reorder(input, output);

        TestHelper.GetPageCount(output).Should().Be(2);
        GetPageWidth(output, 0).Should().Be(101);
        GetPageWidth(output, 1).Should().Be(102);
    }

    [Fact]
    public void Reorder_EvenPageCount_InterleavesCorrectly()
    {
        var input = CreateIdentifiablePdf(4);
        var output = Path.Combine(_tempDir, "reordered.pdf");

        _sut.Reorder(input, output);

        TestHelper.GetPageCount(output).Should().Be(4);
        GetPageWidth(output, 0).Should().Be(101);
        GetPageWidth(output, 1).Should().Be(104);
        GetPageWidth(output, 2).Should().Be(102);
        GetPageWidth(output, 3).Should().Be(103);
    }

    [Fact]
    public void Reorder_OddPageCount_InterleavesCorrectly()
    {
        var input = CreateIdentifiablePdf(5);
        var output = Path.Combine(_tempDir, "reordered.pdf");

        _sut.Reorder(input, output);

        TestHelper.GetPageCount(output).Should().Be(5);
        GetPageWidth(output, 0).Should().Be(101);
        GetPageWidth(output, 1).Should().Be(105);
        GetPageWidth(output, 2).Should().Be(102);
        GetPageWidth(output, 3).Should().Be(104);
        GetPageWidth(output, 4).Should().Be(103);
    }
}
