using System;
using System.IO;
using System.Threading.Tasks;
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

    public void Dispose()
    {
        if (Directory.Exists(_tempDir))
            Directory.Delete(_tempDir, true);
    }

    /// <summary>
    /// Creates a PDF where each page has a different width so we can identify page order.
    /// Page N gets width = 100 + N.
    /// </summary>
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
        doc.Close();
        return path;
    }

    private static double GetPageWidth(string path, int pageIndex)
    {
        using var doc = PdfReader.Open(path, PdfDocumentOpenMode.Modify);
        var width = doc.Pages[pageIndex].Width.Point;
        doc.Close();
        return width;
    }

    [Fact]
    public async Task ReorderAsync_EvenPageCount_InterleavesCorrectly()
    {
        // 4 pages: [1,2,3,4] -> scanned as [1,3] then [4,2] -> input [1,3,4,2]
        // Reorder should produce [1,4,3,2] — first half interleaved with reversed second half
        // Actually: pages [1,2,3,4], middle=2, even
        // i=0: add page[0]=1, add page[3]=4
        // i=1: add page[1]=2, add page[2]=3 (i==middle-1 && isEven)
        // Result: [1,4,2,3]
        var input = CreateIdentifiablePdf(4);
        var output = Path.Combine(_tempDir, "reordered.pdf");

        await PdfService.ReorderAsync(input, output);

        TestHelper.GetPageCount(output).Should().Be(4);
        GetPageWidth(output, 0).Should().Be(101); // page 1
        GetPageWidth(output, 1).Should().Be(104); // page 4
        GetPageWidth(output, 2).Should().Be(102); // page 2
        GetPageWidth(output, 3).Should().Be(103); // page 3
    }

    [Fact]
    public async Task ReorderAsync_OddPageCount_InterleavesCorrectly()
    {
        // 5 pages, middle=3, not even
        // i=0: add page[0]=1, add page[4]=5 (i < middle-1)
        // i=1: add page[1]=2, add page[3]=4 (i < middle-1)
        // i=2: add page[2]=3 (i==middle-1 && !isEven, skip second add)
        // Result: [1,5,2,4,3]
        var input = CreateIdentifiablePdf(5);
        var output = Path.Combine(_tempDir, "reordered.pdf");

        await PdfService.ReorderAsync(input, output);

        TestHelper.GetPageCount(output).Should().Be(5);
        GetPageWidth(output, 0).Should().Be(101); // page 1
        GetPageWidth(output, 1).Should().Be(105); // page 5
        GetPageWidth(output, 2).Should().Be(102); // page 2
        GetPageWidth(output, 3).Should().Be(104); // page 4
        GetPageWidth(output, 4).Should().Be(103); // page 3
    }
}
