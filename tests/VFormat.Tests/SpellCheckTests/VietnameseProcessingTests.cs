using System.Collections.Generic;
using Xunit;
using VFormat.Core.SpellCheck.Processing;
using VFormat.Core.SpellCheck.Models;

namespace VFormat.Tests.SpellCheckTests;

public class VietnameseProcessingTests
{
    [Theory]
    [InlineData('a', true)]
    [InlineData('Z', true)]
    [InlineData('đ', true)]
    [InlineData('Đ', true)]
    [InlineData('á', true)]
    [InlineData('ạ', true)]
    [InlineData('ở', true)]
    [InlineData('ỹ', true)]
    [InlineData('1', false)]
    [InlineData('.', false)]
    [InlineData('@', false)]
    public void IsVietnameseLetter_CategorizesAccurately(char c, bool expected)
    {
        bool actual = VietnameseTokenizer.IsVietnameseLetter(c);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void NormalizeQuotes_ConvertsSmartSingleQuotesToStandardSingleQuote()
    {
        string input = "‘Văn bản hành chính’";
        string normalized = VietnameseUnicodeNormalizer.NormalizeQuotes(input);

        Assert.Equal("'Văn bản hành chính'", normalized);
    }

    [Fact]
    public void NormalizeNfc_ComposedNormalization_ProducesValidNFC()
    {
        // Decomposed 'e' + combining acute ('e' + '\u0301') -> 'é' ('\u00E9')
        string decomposed = "e\u0301";
        string nfc = VietnameseUnicodeNormalizer.NormalizeNfc(decomposed);

        Assert.Equal("é", nfc);
        Assert.True(nfc.IsNormalized(System.Text.NormalizationForm.FormC));
    }

    [Theory]
    [InlineData("hòa", "hoà")]
    [InlineData("hoà", "hòa")]
    [InlineData("thủy", "thuỷ")]
    [InlineData("thuỷ", "thủy")]
    public void GetAlternativeTonePlacement_SwapsModernAndTraditionalTones(string input, string expectedAlternative)
    {
        string alt = VietnameseUnicodeNormalizer.GetAlternativeTonePlacement(input);
        Assert.Equal(expectedAlternative, alt);
    }

    [Fact]
    public void VietnameseTokenizer_TokenizeSimpleSentence_ProducesExpectedTokens()
    {
        var tokenizer = new VietnameseTokenizer();
        string text = "Ủy ban nhân dân xã, ngày 23/10/2026.";

        List<TextToken> tokens = tokenizer.Tokenize(text);

        Assert.NotEmpty(tokens);
        // "Ủy ban nhân dân" is recognized as a protected administrative phrase
        Assert.Contains(tokens, t => t.Type == TokenType.Protected && t.Text == "Ủy ban nhân dân");
        Assert.Contains(tokens, t => t.Type == TokenType.Word && t.Text == "xã");
        Assert.Contains(tokens, t => t.Type == TokenType.Punctuation && t.Text == ",");
        Assert.Contains(tokens, t => t.Type == TokenType.Whitespace);
    }
}
