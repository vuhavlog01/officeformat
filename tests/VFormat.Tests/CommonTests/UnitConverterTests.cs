using System;
using Xunit;
using VFormat.Core.Common;

namespace VFormat.Tests.CommonTests;

public class UnitConverterTests
{
    private const float PrecisionDelta = 0.001f;

    [Theory]
    [InlineData(0f, 0f)]
    [InlineData(10f, 28.346f)]
    [InlineData(20f, 56.693f)]
    [InlineData(25f, 70.866f)]
    [InlineData(30f, 85.039f)]
    [InlineData(75f, 212.598f)]
    [InlineData(95f, 269.291f)]
    public void MmToPt_FloatValues_ConvertsCorrectly(float mm, float expectedPt)
    {
        float actual = UnitConverter.MmToPt(mm);
        Assert.InRange(actual, expectedPt - 0.05f, expectedPt + 0.05f);
    }

    [Theory]
    [InlineData(0.0, 0.0)]
    [InlineData(20.0, 56.6929)]
    [InlineData(30.0, 85.0393)]
    public void MmToPt_DoubleValues_ConvertsCorrectly(double mm, double expectedPt)
    {
        double actual = UnitConverter.MmToPt(mm);
        Assert.InRange(actual, expectedPt - 0.01, expectedPt + 0.01);
    }

    [Theory]
    [InlineData(0f, 0f)]
    [InlineData(28.346457f, 10f)]
    [InlineData(56.692914f, 20f)]
    [InlineData(85.039371f, 30f)]
    public void PtToMm_FloatValues_ConvertsCorrectly(float pt, float expectedMm)
    {
        float actual = UnitConverter.PtToMm(pt);
        Assert.InRange(actual, expectedMm - 0.05f, expectedMm + 0.05f);
    }

    [Fact]
    public void MmAndPt_RoundTrip_PreservesValue()
    {
        float originalMm = 25f;
        float pt = UnitConverter.MmToPt(originalMm);
        float roundtripMm = UnitConverter.PtToMm(pt);

        Assert.Equal(originalMm, roundtripMm, 2);
    }

    [Theory]
    [InlineData(1f, 28.346f)]
    [InlineData(2.5f, 70.866f)]
    public void CmToPt_FloatValues_ConvertsCorrectly(float cm, float expectedPt)
    {
        float actual = UnitConverter.CmToPt(cm);
        Assert.InRange(actual, expectedPt - 0.05f, expectedPt + 0.05f);
    }

    [Fact]
    public void PtToTwips_TwipsToPt_RoundTrip()
    {
        float originalPt = 12f;
        int twips = UnitConverter.PtToTwips(originalPt);
        Assert.Equal(240, twips);

        float roundtripPt = UnitConverter.TwipsToPt(twips);
        Assert.Equal(originalPt, roundtripPt);
    }

    [Fact]
    public void MmToTwips_ConvertsCorrectly()
    {
        float mm = 10f;
        int twips = UnitConverter.MmToTwips(mm);
        Assert.Equal(567, twips);
    }

    [Theory]
    [InlineData(12.1f, 12.0f)]
    [InlineData(12.26f, 12.5f)]
    [InlineData(12.74f, 12.5f)]
    [InlineData(12.8f, 13.0f)]
    [InlineData(14.0f, 14.0f)]
    public void RoundToHalfPoint_RoundsProperly(float input, float expected)
    {
        float actual = UnitConverter.RoundToHalfPoint(input);
        Assert.Equal(expected, actual);
    }
}
