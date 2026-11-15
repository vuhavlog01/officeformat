using System.Text.RegularExpressions;
using Xunit;
using VFormat.Core.Common;

namespace VFormat.Tests.CommonTests;

public class RegexPatternsTests
{
    [Theory]
    [InlineData("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", true)]
    [InlineData("CỘNG  HÒA   XÃ HỘI CHỦ NGHĨA   VIỆT NAM", true)]
    [InlineData("ĐẢNG CỘNG SẢN VIỆT NAM", false)]
    [InlineData("BỘ NỘI VỤ", false)]
    public void NationalHeader_MatchesOnlyNationalHeader(string input, bool expected)
    {
        bool actual = RegexPatterns.NationalHeader.IsMatch(input);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("Độc lập - Tự do - Hạnh phúc", true)]
    [InlineData("Độc lập – Tự do – Hạnh phúc", true)] // en-dash
    [InlineData("Độc lập — Tự do — Hạnh phúc", true)] // em-dash
    [InlineData("Độc lập - Tự do", false)]
    [InlineData("Hà Nội, ngày 20 tháng 10 năm 2026", false)]
    public void Motto_MatchesMottoVariations(string input, bool expected)
    {
        bool actual = RegexPatterns.Motto.IsMatch(input);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("ĐẢNG CỘNG SẢN VIỆT NAM", true)]
    [InlineData("ĐẢNG  CỘNG  SẢN  VIỆT  NAM", true)]
    [InlineData("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", false)]
    public void PartyHeader_MatchesPartyHeader(string input, bool expected)
    {
        bool actual = RegexPatterns.PartyHeader.IsMatch(input);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("ĐOÀN TNCS HỒ CHÍ MINH", true)]
    [InlineData("ĐOÀN THANH NIÊN CỘNG SẢN HỒ CHÍ MINH", true)]
    [InlineData("BCH ĐOÀN TỈNH BẮC GIANG", false)]
    public void YouthUnionHeader_MatchesYouthHeader(string input, bool expected)
    {
        bool actual = RegexPatterns.YouthUnionHeader.IsMatch(input);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("Số: 123/UBND-TH", true)]
    [InlineData("Số: 01/2026/QĐ-UBND", true)]
    [InlineData("Số 45-KH/TU", true)]
    [InlineData("Số: 12/HD-VPTW", true)]
    [InlineData("Số 245-CV/TWĐTN-VP", true)]
    [InlineData("Hà Nội, ngày 23 tháng 10 năm 2026", false)]
    public void DocumentNumberGeneral_MatchesDocumentNumbers(string input, bool expected)
    {
        bool actual = RegexPatterns.DocumentNumberGeneral.IsMatch(input);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("Hà Nội, ngày 23 tháng 10 năm 2026", true)]
    [InlineData("TP. Hồ Chí Minh, ngày 15 tháng 8 năm 2025", true)]
    [InlineData("..., ngày ... tháng ... năm 20...", false)] // ellipsis is not \w
    [InlineData("Số: 123/UBND", false)]
    public void DocumentDate_MatchesDocumentDates(string input, bool expected)
    {
        bool actual = RegexPatterns.DocumentDate.IsMatch(input);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("TM.", true)]
    [InlineData("KT.", true)]
    [InlineData("TL.", true)]
    [InlineData("TUQ.", true)]
    [InlineData("Q. CHỦ TỊCH", true)]
    [InlineData("T/M", true)]
    [InlineData("K/T", true)]
    [InlineData("T/L", true)]
    [InlineData("CHỦ TỊCH", false)]
    public void SignAuthority_MatchesSignAuthorityPrefixes(string input, bool expected)
    {
        bool actual = RegexPatterns.SignAuthority.IsMatch(input);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("Nơi nhận:", true)]
    [InlineData("Nơi nhận :", true)]
    [InlineData("NƠI NHẬN:", true)]
    [InlineData("- Như trên;", false)]
    public void RecipientHeader_MatchesRecipientHeader(string input, bool expected)
    {
        bool actual = RegexPatterns.RecipientHeader.IsMatch(input);
        Assert.Equal(expected, actual);
    }
}
