using FluentAssertions;
using PdfChopper.Models;
using Xunit;

namespace PdfChopper.Tests.Models;

public class PdfFileRotationTests
{
    [Fact]
    public void Rotate_DefaultsToZero()
    {
        var rotation = new PdfFileRotation(5);

        rotation.Rotate.Should().Be(0);
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 1)]
    [InlineData(2, 2)]
    [InlineData(3, 3)]
    public void Rotate_AcceptsValuesZeroToThree(int input, int expected)
    {
        var rotation = new PdfFileRotation(5);

        rotation.Rotate = input;

        rotation.Rotate.Should().Be(expected);
    }

    [Theory]
    [InlineData(4, 0)]
    [InlineData(5, 1)]
    [InlineData(7, 3)]
    [InlineData(8, 0)]
    public void Rotate_NormalizesLargeValues(int input, int expected)
    {
        var rotation = new PdfFileRotation(5);

        rotation.Rotate = input;

        rotation.Rotate.Should().Be(expected);
    }

    [Theory]
    [InlineData(-1, 3)]
    [InlineData(-2, 2)]
    [InlineData(-3, 1)]
    [InlineData(-4, 0)]
    [InlineData(-5, 3)]
    public void Rotate_NormalizesNegativeValues(int input, int expected)
    {
        var rotation = new PdfFileRotation(5);

        rotation.Rotate = input;

        rotation.Rotate.Should().Be(expected);
    }
}
