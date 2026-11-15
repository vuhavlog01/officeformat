using System;
using Xunit;
using VFormat.Core.Common;
using VFormat.Core.Enums;
using VFormat.Core.Models;
using VFormat.Core.Rules;
using VFormat.WordEngine;

namespace VFormat.Tests.RulesTests;

public class SelectionFormattingTests
{
    [Fact]
    public void StandardizeSelection_NullSelection_ReturnsFailed()
    {
        var engine = new WordFormatEngine();
        var result = engine.StandardizeSelection(null);
        Assert.False(result.Success);
        Assert.NotNull(result.Error);
        Assert.Contains("không hợp lệ", result.Message);
    }

    [Fact]
    public void FormattingResult_Succeeded_SetsParagraphsAndTables()
    {
        var result = FormattingResult.Succeeded(RulePackType.ND30_2020, 5, 2, 120);
        Assert.True(result.Success);
        Assert.Equal(5, result.ParagraphsProcessed);
        Assert.Equal(2, result.TablesProcessed);
        Assert.Equal(RulePackType.ND30_2020, result.RulePack);
        Assert.Contains("5 đoạn văn", result.Message);
        Assert.Contains("2 bảng", result.Message);
    }

    [Theory]
    [InlineData(RulePackType.ND30_2020, SemanticRole.DocumentType, true, false)]
    [InlineData(RulePackType.ND30_2020, SemanticRole.SubjectExcerpt, true, false)]
    [InlineData(RulePackType.ND30_2020, SemanticRole.LegalGround, false, true)]
    [InlineData(RulePackType.HD05_VPTW, SemanticRole.DocumentType, true, false)]
    [InlineData(RulePackType.HD05_VPTW, SemanticRole.SubjectExcerpt, true, false)]
    public void DocumentStyleResolver_ResolvesSelectedStylesCorrectly(RulePackType pack, SemanticRole role, bool expectedBold, bool expectedItalic)
    {
        var engine = new RuleEngine();
        engine.SetActiveProfile(pack);
        var profile = engine.ActiveProfile;

        var style = DocumentStyleResolver.Resolve(profile, role, "QUYẾT ĐỊNH");
        Assert.NotNull(style);
        if (role == SemanticRole.LegalGround)
        {
            Assert.Equal(expectedItalic, style.Italic);
            Assert.Equal(expectedBold, style.Bold);
        }
        else
        {
            Assert.Equal(expectedBold, style.Bold);
        }
    }

    [Theory]
    [InlineData("I. MỤC ĐÍCH, YÊU CẦU", true)]
    [InlineData("II. NỘI DUNG THỰC HIỆN", true)]
    [InlineData("IV. TỔ CHỨC THỰC HIỆN", true)]
    [InlineData("1. Nhiệm vụ thường xuyên", false)]
    [InlineData("Điều 1. Phạm vi điều chỉnh", false)]
    public void WordFormatEngine_IsRomanNumeralHeading_DetectsCorrectly(string text, bool expected)
    {
        bool actual = WordFormatEngine.IsRomanNumeralHeading(text);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("1. Công tác chỉ đạo điều hành:", true)]
    [InlineData("2. Nhiệm vụ trọng tâm năm 2026", true)]
    [InlineData("12. Các nội dung khác", true)]
    [InlineData("Hôm nay ngày 20 tháng 10 năm 2026", false)]
    public void WordFormatEngine_IsClauseHeading_DetectsCorrectly(string text, bool expected)
    {
        bool actual = WordFormatEngine.IsClauseHeading(text);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("1. Về công tác cán bộ:", true)]
    [InlineData("2. Nhiệm vụ cụ thể:", true)]
    [InlineData("Một đoạn văn dài bình thường trong nội dung thân bài không phải là điều khoản có tiêu đề ngắn.", false)]
    public void WordFormatEngine_IsTitledClause_DetectsCorrectly(string text, bool expected)
    {
        bool actual = WordFormatEngine.IsTitledClause(text);
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("Điều 1. Phạm vi điều chỉnh", true)]
    [InlineData("Điều 2. Đối tượng áp dụng", true)]
    [InlineData("Không phải điều khoản", false)]
    public void RegexPatterns_ArticleTitle_DetectsCorrectly(string text, bool expected)
    {
        bool actual = RegexPatterns.ArticleTitle.IsMatch(text);
        Assert.Equal(expected, actual);
    }

    [Fact]
    public void DeleteBlankPages_NullDocument_ReturnsZero()
    {
        int result = WordToolEngine.DeleteBlankPages(null);
        Assert.Equal(0, result);
    }
}
