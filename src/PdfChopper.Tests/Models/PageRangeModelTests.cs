using FluentAssertions;
using PdfChopper.Models;
using Xunit;

namespace PdfChopper.Tests.Models;

public class PageRangeModelTests
{
    private static PdfFile CreateModel(int pageCount) => new("test.pdf", pageCount);

    [Fact]
    public void Constructor_SetsDefaultRange()
    {
        var model = CreateModel(5);

        model.StartPage.Should().Be(1);
        model.EndPage.Should().Be(5);
        model.PageCount.Should().Be(5);
    }

    [Fact]
    public void StartPage_AcceptsValidValue()
    {
        var model = CreateModel(5);

        model.StartPage = 3;

        model.StartPage.Should().Be(3);
    }

    [Fact]
    public void StartPage_RejectsZero()
    {
        var model = CreateModel(5);

        model.StartPage = 0;

        model.StartPage.Should().Be(1);
    }

    [Fact]
    public void StartPage_RejectsNegative()
    {
        var model = CreateModel(5);

        model.StartPage = -1;

        model.StartPage.Should().Be(1);
    }

    [Fact]
    public void StartPage_RejectsValueAbovePageCount()
    {
        var model = CreateModel(5);

        model.StartPage = 6;

        model.StartPage.Should().Be(1);
    }

    [Fact]
    public void StartPage_RejectsValueAboveEndPage()
    {
        var model = CreateModel(5);
        model.EndPage = 3;

        model.StartPage = 4;

        model.StartPage.Should().Be(1);
    }

    [Fact]
    public void StartPage_AcceptsValueEqualToEndPage()
    {
        var model = CreateModel(5);
        model.EndPage = 3;

        model.StartPage = 3;

        model.StartPage.Should().Be(3);
    }

    [Fact]
    public void EndPage_AcceptsValidValue()
    {
        var model = CreateModel(5);

        model.EndPage = 3;

        model.EndPage.Should().Be(3);
    }

    [Fact]
    public void EndPage_RejectsZero()
    {
        var model = CreateModel(5);

        model.EndPage = 0;

        model.EndPage.Should().Be(5);
    }

    [Fact]
    public void EndPage_RejectsNegative()
    {
        var model = CreateModel(5);

        model.EndPage = -1;

        model.EndPage.Should().Be(5);
    }

    [Fact]
    public void EndPage_RejectsValueAbovePageCount()
    {
        var model = CreateModel(5);

        model.EndPage = 6;

        model.EndPage.Should().Be(5);
    }

    [Fact]
    public void EndPage_RejectsValueBelowStartPage()
    {
        var model = CreateModel(5);
        model.StartPage = 3;

        model.EndPage = 2;

        model.EndPage.Should().Be(5);
    }

    [Fact]
    public void EndPage_AcceptsValueEqualToStartPage()
    {
        var model = CreateModel(5);
        model.StartPage = 3;

        model.EndPage = 3;

        model.EndPage.Should().Be(3);
    }

    [Fact]
    public void StartAndEnd_CanNarrowToSinglePage()
    {
        var model = CreateModel(10);

        model.EndPage = 5;
        model.StartPage = 5;

        model.StartPage.Should().Be(5);
        model.EndPage.Should().Be(5);
    }
}
