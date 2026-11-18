using System.IO;
using Xunit;
using VFormat.Core.Personalization;
using VFormat.Core.Serialization;
using VFormat.Core.Enums;

namespace VFormat.Tests.PersonalizationTests;

public class ProfileSerializationTests
{
    [Theory]
    [InlineData("2.5", 2.5f)]
    [InlineData("2,5", 2.5f)]
    [InlineData("  14.25  ", 14.25f)]
    [InlineData("14,25", 14.25f)]
    [InlineData("0", 0f)]
    [InlineData("100", 100f)]
    public void TryParseFlexibleFloat_ValidInputs_ParsesAccurately(string input, float expected)
    {
        bool success = JsonProfileRepository.TryParseFlexibleFloat(input, out float actual);
        Assert.True(success, $"Failed for {input}");
        Assert.Equal(expected, actual, 2);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(null)]
    [InlineData("abc")]
    [InlineData("12.34.56")]
    public void TryParseFlexibleFloat_InvalidInputs_ReturnsFalse(string invalidInput)
    {
        bool success = JsonProfileRepository.TryParseFlexibleFloat(invalidInput, out _);
        Assert.False(success);
    }

    [Fact]
    public void UserPersonalization_InstantiateDefaults_HasValidCollections()
    {
        var p = new UserPersonalization();

        Assert.NotNull(p.AgenciesAdmin);
        Assert.NotNull(p.AgenciesParty);
        Assert.NotNull(p.AgenciesMilitary);
        Assert.NotNull(p.AgenciesYouth);
        Assert.NotNull(p.SignersAdmin);
        Assert.NotNull(p.SignersParty);
        Assert.NotNull(p.SignersYouth);
    }

    [Fact]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Assertions", "xUnit2000:Constants should be on the left")]
    public void MarginProfile_AssignAndRetrieve_HoldsValues()
    {
        var margin = new MarginProfile
        {
            MarginTopMm = 20f,
            MarginBottomMm = 20f,
            MarginLeftMm = 30f,
            MarginRightMm = 15f
        };

        Assert.Equal(20f, margin.MarginTopMm);
        Assert.Equal(20f, margin.MarginBottomMm);
        Assert.Equal(30f, margin.MarginLeftMm);
        Assert.Equal(15f, margin.MarginRightMm);
    }

    [Fact]
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Assertions", "xUnit2000:Constants should be on the left")]
    public void TypographyProfile_AssignAndRetrieve_HoldsValues()
    {
        var typo = new TypographyProfile
        {
            SizeQuocHieu = 12f,
            SizeThanBai = 14f,
            BoldChuKy = true
        };

        Assert.Equal(12f, typo.SizeQuocHieu);
        Assert.Equal(14f, typo.SizeThanBai);
        Assert.True(typo.BoldChuKy);
    }

    [Fact]
    public void JsonProfileRepository_LoadAndSave_RoundtripsSuccessfully()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), "vformat_test_presets_" + System.Guid.NewGuid().ToString("N") + ".json");
        try
        {
            var repo = new JsonProfileRepository(tempFile);
            var personalization = repo.LoadPersonalization();
            Assert.NotNull(personalization);

            // Modify a value
            personalization.HeaderDateInTopCell = true;
            personalization.ActiveRulePack = RulePackType.HD05_VPTW;
            repo.SavePersonalization(personalization);

            // Re-read with a fresh repository instance
            var repo2 = new JsonProfileRepository(tempFile);
            var loaded = repo2.LoadPersonalization();

            Assert.True(loaded.HeaderDateInTopCell);
            Assert.Equal(RulePackType.HD05_VPTW, loaded.ActiveRulePack);
        }
        finally
        {
            if (File.Exists(tempFile)) File.Delete(tempFile);
            if (File.Exists(tempFile + ".bak")) File.Delete(tempFile + ".bak");
        }
    }

    [Fact]
    public void MarginProfile_HD05Default_MatchesWordParagraphSpecification()
    {
        var hd05 = MarginProfile.CreateHD05Default();

        Assert.Equal(20f, hd05.MarginTopMm);
        Assert.Equal(20f, hd05.MarginBottomMm);
        Assert.Equal(30f, hd05.MarginLeftMm);
        Assert.Equal(15f, hd05.MarginRightMm);
        Assert.Equal(1f, hd05.FirstLineIndentCm);
        Assert.Equal(6f, hd05.SpacingBeforePt);
        Assert.Equal(6f, hd05.SpacingAfterPt);
        Assert.Equal("Exactly", hd05.LineSpacingRule);
        Assert.Equal(18f, hd05.LineSpacingAtPt);
        Assert.False(hd05.NoSpaceBetweenParagraphsOfSameStyle);
    }

    [Fact]
    public void JsonProfileRepository_HD05MarginsAndSpacing_RoundtripsSuccessfully()
    {
        string tempFile = Path.Combine(Path.GetTempPath(), "vformat_test_hd05_" + System.Guid.NewGuid().ToString("N") + ".json");
        try
        {
            var repo = new JsonProfileRepository(tempFile);
            var personalization = UserPersonalization.CreateDefault();
            repo.SavePersonalization(personalization);

            var loadedInit = repo.LoadPersonalization();
            Assert.NotNull(loadedInit);
            Assert.NotNull(loadedInit.MarginsParty);
            Assert.Equal(6f, loadedInit.MarginsParty.SpacingBeforePt);
            Assert.Equal(6f, loadedInit.MarginsParty.SpacingAfterPt);
            Assert.Equal("Exactly", loadedInit.MarginsParty.LineSpacingRule);
            Assert.Equal(18f, loadedInit.MarginsParty.LineSpacingAtPt);
            Assert.False(loadedInit.MarginsParty.NoSpaceBetweenParagraphsOfSameStyle);

            // Modify values
            personalization.MarginsParty.SpacingBeforePt = 12f;
            personalization.MarginsParty.SpacingAfterPt = 8f;
            personalization.MarginsParty.LineSpacingAtPt = 20f;
            personalization.MarginsParty.NoSpaceBetweenParagraphsOfSameStyle = true;
            repo.SavePersonalization(personalization);

            // Re-read with a fresh repository instance
            var repo2 = new JsonProfileRepository(tempFile);
            var loaded = repo2.LoadPersonalization();

            Assert.Equal(12f, loaded.MarginsParty.SpacingBeforePt);
            Assert.Equal(8f, loaded.MarginsParty.SpacingAfterPt);
            Assert.Equal(20f, loaded.MarginsParty.LineSpacingAtPt);
            Assert.True(loaded.MarginsParty.NoSpaceBetweenParagraphsOfSameStyle);
        }
        finally
        {
            if (File.Exists(tempFile)) File.Delete(tempFile);
            if (File.Exists(tempFile + ".bak")) File.Delete(tempFile + ".bak");
        }
    }

    [Fact]
    public void HD05Rules_CreateProfile_HasSpacingBefore6AndExactly18()
    {
        var profile = VFormat.Core.Rules.HD05Rules.CreateProfile();

        Assert.Equal(6f, profile.Typography.SpacingBeforePt);
        Assert.Equal(6f, profile.Typography.SpacingAfterPt);
        Assert.Equal(LineSpacingType.Exactly, profile.Typography.LineSpacingRule);
        Assert.Equal(18f, profile.Typography.LineSpacingValue);
        Assert.False(profile.Typography.NoSpaceBetweenParagraphsOfSameStyle);

        var bodyStyle = profile.GetComponentStyle(SemanticRole.BodyText);
        Assert.NotNull(bodyStyle);
        Assert.Equal(6f, bodyStyle.SpacingBeforePt);
        Assert.Equal(6f, bodyStyle.SpacingAfterPt);
        Assert.Equal(LineSpacingType.Exactly, bodyStyle.LineSpacingRule);
        Assert.Equal(18f, bodyStyle.LineSpacingValue);
    }

    [Fact]
    public void RuleEngine_HD05_AppliesPersonalizationToActiveProfile()
    {
        var ruleEngine = new VFormat.Core.Rules.RuleEngine(null, UserPersonalization.CreateDefault());
        ruleEngine.SetActiveProfile(RulePackType.HD05_VPTW);

        var active = ruleEngine.ActiveProfile;
        Assert.Equal(6f, active.Typography.SpacingBeforePt);
        Assert.Equal(6f, active.Typography.SpacingAfterPt);
        Assert.Equal(LineSpacingType.Exactly, active.Typography.LineSpacingRule);
        Assert.Equal(18f, active.Typography.LineSpacingValue);

        var body = active.GetComponentStyle(SemanticRole.BodyText);
        Assert.Equal(6f, body.SpacingBeforePt);
        Assert.Equal(6f, body.SpacingAfterPt);
        Assert.Equal(LineSpacingType.Exactly, body.LineSpacingRule);
        Assert.Equal(18f, body.LineSpacingValue);
    }
}
