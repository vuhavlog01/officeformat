using Xunit;
using VFormat.Core.Common;

namespace VFormat.Tests.CommonTests;

public class VndAmountReaderTests
{
    [Theory]
    [InlineData("0", "Không đồng")]
    [InlineData("1", "Một đồng")]
    [InlineData("5", "Năm đồng")]
    [InlineData("10", "Mười đồng")]
    [InlineData("15", "Mười lăm đồng")]
    [InlineData("21", "Hai mươi mốt đồng")]
    [InlineData("25", "Hai mươi lăm đồng")]
    [InlineData("105", "Một trăm linh năm đồng")]
    [InlineData("115", "Một trăm mười lăm đồng")]
    [InlineData("1000", "Một nghìn đồng")]
    [InlineData("15000", "Mười lăm nghìn đồng")]
    [InlineData("1000000", "Một triệu đồng")]
    [InlineData("1000000000", "Một tỷ đồng")]
    public void TryConvert_StandardNumbers_ReadsCorrectly(string input, string expected)
    {
        bool success = VndAmountReader.TryConvert(input, out string words, out string error);
        Assert.True(success, $"Failed for {input}: {error}");
        Assert.Equal(expected, words);
    }

    [Theory]
    [InlineData("50.000 vnđ", "Năm mươi nghìn đồng")]
    [InlineData("50.000đ", "Năm mươi nghìn đồng")]
    [InlineData("50.000 đồng", "Năm mươi nghìn đồng")]
    [InlineData("50.000 ₫", "Năm mươi nghìn đồng")]
    [InlineData("50.000 VND", "Năm mươi nghìn đồng")]
    public void TryConvert_WithCurrencySuffixes_ReadsCorrectly(string input, string expected)
    {
        bool success = VndAmountReader.TryConvert(input, out string words, out string error);
        Assert.True(success, $"Failed for {input}: {error}");
        Assert.Equal(expected, words);
    }

    [Theory]
    [InlineData("1.234.567", "Một triệu hai trăm ba mươi bốn nghìn năm trăm sáu mươi bảy đồng")]
    [InlineData("1,234,567", "Một triệu hai trăm ba mươi bốn nghìn năm trăm sáu mươi bảy đồng")]
    [InlineData("1 234 567", "Một triệu hai trăm ba mươi bốn nghìn năm trăm sáu mươi bảy đồng")]
    [InlineData("1234567", "Một triệu hai trăm ba mươi bốn nghìn năm trăm sáu mươi bảy đồng")]
    public void TryConvert_DifferentGroupingSeparators_ReadsEqually(string input, string expected)
    {
        bool success = VndAmountReader.TryConvert(input, out string words, out string error);
        Assert.True(success, $"Failed for {input}: {error}");
        Assert.Equal(expected, words);
    }

    [Fact]
    public void TryConvert_NegativeNumber_IncludesAmPrefix()
    {
        bool success = VndAmountReader.TryConvert("-50000", out string words, out string error);
        Assert.True(success, $"Failed: {error}");
        Assert.StartsWith("Âm", words);
        Assert.Contains("năm mươi nghìn đồng", words);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("12.34.56.78")]
    [InlineData("10.000,50")] // Non-zero decimal in VND
    public void TryConvert_InvalidOrDecimalInputs_ReturnsFalse(string invalidInput)
    {
        bool success = VndAmountReader.TryConvert(invalidInput, out string words, out string error);
        Assert.False(success);
        Assert.NotNull(error);
        Assert.Null(words);
    }

    [Fact]
    public void VndReadingOptions_CustomOptions_AppliesVariations()
    {
        var options = new VndReadingOptions
        {
            ThousandAsNgan = true,
            FourAsTu = true,
            ZeroTensAsLe = true,
            CapitalizeFirst = true
        };

        bool success = VndAmountReader.TryConvert("104", options, out string words1, out _);
        Assert.True(success);
        Assert.Equal("Một trăm lẻ bốn đồng", words1);

        bool success2 = VndAmountReader.TryConvert("1024", options, out string words2, out _);
        Assert.True(success2);
        Assert.Contains("ngàn", words2);
        Assert.Contains("tư", words2);
    }
}
