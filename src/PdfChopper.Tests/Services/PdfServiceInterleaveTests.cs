using System;
using System.Collections.Generic;
using System.IO;
using FluentAssertions;
using PdfChopper.Models;
using PdfChopper.Services;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using Xunit;

namespace PdfChopper.Tests.Services;

public class PdfServiceInterleaveTests : IDisposable
{
    private readonly string _tempDir = TestHelper.CreateTempDir();
    private readonly PdfService _sut = new();

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    private string CreateIdentifiablePdf(int pageCount, int widthBase)
    {
        var path = Path.Combine(_tempDir, $"{Guid.NewGuid()}.pdf");
        using var doc = new PdfDocument();
        for (var i = 1; i <= pageCount; i++)
        {
            doc.AddPage(new PdfPage { Width = XUnit.FromPoint(widthBase + i) });
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
    public void Interleave_TwoEqualFiles_ProducesCorrectPageOrder()
    {
        var file1 = CreateIdentifiablePdf(2, 100);
        var file2 = CreateIdentifiablePdf(2, 200);
        var output = Path.Combine(_tempDir, "interleaved.pdf");

        var files = new List<PdfFile> { new(file1, 2), new(file2, 2) };

        _sut.Interleave(files, output);

        TestHelper.GetPageCount(output).Should().Be(4);
        GetPageWidth(output, 0).Should().Be(101);
        GetPageWidth(output, 1).Should().Be(201);
        GetPageWidth(output, 2).Should().Be(102);
        GetPageWidth(output, 3).Should().Be(202);
    }

    [Fact]
    public void Interleave_UnequalPageCounts_IncludesAllPagesInOrder()
    {
        var file1 = CreateIdentifiablePdf(3, 100);
        var file2 = CreateIdentifiablePdf(1, 200);
        var output = Path.Combine(_tempDir, "interleaved.pdf");

        var files = new List<PdfFile> { new(file1, 3), new(file2, 1) };

        _sut.Interleave(files, output);

        TestHelper.GetPageCount(output).Should().Be(4);
        GetPageWidth(output, 0).Should().Be(101);
        GetPageWidth(output, 1).Should().Be(201);
        GetPageWidth(output, 2).Should().Be(102);
        GetPageWidth(output, 3).Should().Be(103);
    }
}
