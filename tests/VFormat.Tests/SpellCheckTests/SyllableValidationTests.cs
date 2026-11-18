using Xunit;
using VFormat.Core.SpellCheck.Validation;

namespace VFormat.Tests.SpellCheckTests;

public class SyllableValidationTests
{
    [Theory]
    [InlineData("nghiêng")]
    [InlineData("hoàng")]
    [InlineData("thỏa")]
    [InlineData("việt")]
    [InlineData("nam")]
    [InlineData("quốc")]
    public void Validate_ValidVietnameseSyllables_ReturnsValid(string word)
    {
        SyllableValidationResult result = SyllableValidator.Validate(word);
        Assert.Equal(SyllableValidationResult.ValidSyllable, result);
    }

    [Theory]
    [InlineData("wifi")]
    [InlineData("file")]
    [InlineData("zalo")]
    [InlineData("java")]
    public void Validate_ForeignWordsWithFJWZ_ReturnsForeignOrUnknown(string word)
    {
        SyllableValidationResult result = SyllableValidator.Validate(word);
        Assert.Equal(SyllableValidationResult.ForeignOrUnknown, result);
    }

    [Theory]
    [InlineData("bcdf")]
    [InlineData("khiênng")]
    [InlineData("")]
    [InlineData(null)]
    public void Validate_InvalidOrEmptyWords_ReturnsInvalid(string word)
    {
        SyllableValidationResult result = SyllableValidator.Validate(word);
        Assert.NotEqual(SyllableValidationResult.ValidSyllable, result);
    }
}
