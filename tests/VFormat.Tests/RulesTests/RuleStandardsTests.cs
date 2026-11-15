using Xunit;
using VFormat.Core.Enums;
using VFormat.Core.Rules;

namespace VFormat.Tests.RulesTests;

public class RuleStandardsTests
{
    [Fact]
    public void RulePackTraits_PartyHeader_HasCorrectConstant()
    {
        Assert.Equal("ĐẢNG CỘNG SẢN VIỆT NAM", RulePackTraits.PartyHeaderText);
        Assert.Equal("ĐOÀN TNCS HỒ CHÍ MINH", RulePackTraits.YouthUnionHeaderText);
    }

    [Theory]
    [InlineData(RulePackType.ND30_2020, false)]
    [InlineData(RulePackType.HD05_VPTW, true)]
    [InlineData(RulePackType.HD29_TWDTN, true)]
    [InlineData(RulePackType.TT68_2023_BQP, false)]
    public void RulePackTraits_UsesOrganizationHeader(RulePackType pack, bool expected)
    {
        bool actual = RulePackTraits.UsesOrganizationHeader(pack);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData(RulePackType.ND30_2020, SemanticRole.NationalHeader, 12, 13)]
    [InlineData(RulePackType.ND30_2020, SemanticRole.Motto, 13, 14)]
    [InlineData(RulePackType.ND30_2020, SemanticRole.DocumentNumber, 13, 13)]
    [InlineData(RulePackType.ND30_2020, SemanticRole.DocumentDate, 13, 14)]
    [InlineData(RulePackType.ND30_2020, SemanticRole.BodyText, 13, 14)]
    [InlineData(RulePackType.ND30_2020, SemanticRole.RecipientHeader, 12, 12)]
    public void LegalFontSizeStandards_ND30_EnforcesCorrectRanges(RulePackType pack, SemanticRole role, int expectedMin, int expectedMax)
    {
        LegalFontSizeRange range = LegalFontSizeStandards.GetLegalRange(pack, role);
        Assert.Equal(expectedMin, range.Min);
        Assert.Equal(expectedMax, range.Max);
    }

    [Theory]
    [InlineData(RulePackType.HD05_VPTW, SemanticRole.PartyHeader, 15, 15)]
    [InlineData(RulePackType.HD05_VPTW, SemanticRole.BodyText, 14, 15)]
    [InlineData(RulePackType.HD05_VPTW, SemanticRole.RecipientHeader, 14, 14)]
    [InlineData(RulePackType.HD05_VPTW, SemanticRole.RecipientItem, 12, 12)]
    public void LegalFontSizeStandards_HD05_EnforcesPartyRanges(RulePackType pack, SemanticRole role, int expectedMin, int expectedMax)
    {
        LegalFontSizeRange range = LegalFontSizeStandards.GetLegalRange(pack, role);
        Assert.Equal(expectedMin, range.Min);
        Assert.Equal(expectedMax, range.Max);
    }

    [Theory]
    [InlineData(RulePackType.ND30_2020, SemanticRole.BodyText, 10f, 13f)] // 10 is below 13, clamped to 13
    [InlineData(RulePackType.ND30_2020, SemanticRole.BodyText, 18f, 14f)] // 18 is above 14, clamped to 14
    [InlineData(RulePackType.ND30_2020, SemanticRole.BodyText, 14f, 14f)] // exactly 14
    public void LegalFontSizeStandards_EnsureLegalInteger_ClampsToRange(RulePackType pack, SemanticRole role, float inputSize, float expectedClamped)
    {
        float actual = LegalFontSizeStandards.EnsureLegalInteger(pack, role, inputSize);
        Assert.Equal(expectedClamped, actual);
    }
}
