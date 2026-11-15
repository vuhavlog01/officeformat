using System.Collections.Generic;
using Xunit;
using VFormat.Core.SpellCheck;
using VFormat.Core.SpellCheck.Models;

namespace VFormat.Tests.SpellCheckTests;

public class SpacingCheckScannerTests
{
    [Fact]
    public void CheckParagraph_DoubleSpaces_DetectsIssue()
    {
        string text = "Cộng hòa  xã hội chủ nghĩa Việt Nam.";
        List<SpellCheckIssue> issues = SpacingCheckScanner.CheckParagraph(text, 1);

        Assert.NotEmpty(issues);
        Assert.Contains(issues, i => i.RuleId == "SPACE_MULTIPLE");
    }

    [Fact]
    public void CheckParagraph_SpaceBeforePunctuation_DetectsIssue()
    {
        string text = "Ủy ban nhân dân , thành phố Hà Nội .";
        List<SpellCheckIssue> issues = SpacingCheckScanner.CheckParagraph(text, 1);

        Assert.NotEmpty(issues);
        Assert.Contains(issues, i => i.RuleId == "SPACE_BEFORE_PUNCT");
    }

    [Fact]
    public void CheckParagraph_StraightQuotes_DetectsIssue()
    {
        string text = "Báo cáo về \"Tình hình kinh tế\".";
        List<SpellCheckIssue> issues = SpacingCheckScanner.CheckParagraph(text, 1);

        Assert.NotEmpty(issues);
        Assert.Contains(issues, i => i.RuleId == "SPACE_STRAIGHT_QUOTE");
    }

    [Fact]
    public void CheckParagraph_CleanText_HasNoIssues()
    {
        string text = "Cộng hòa xã hội chủ nghĩa Việt Nam, ngày 23 tháng 10 năm 2026.";
        List<SpellCheckIssue> issues = SpacingCheckScanner.CheckParagraph(text, 1);

        Assert.Empty(issues);
    }

    [Fact]
    public void CheckParagraph_ProtectedDateBlanks_DoesNotFlagAsDoubleSpace()
    {
        string text = "ngày      tháng      năm 2026";
        List<SpellCheckIssue> issues = SpacingCheckScanner.CheckParagraph(text, 1);

        // ProtectedBlankPattern shields "ngày ... tháng" and "tháng ... năm"
        Assert.DoesNotContain(issues, i => i.RuleId == "SPACE_MULTIPLE");
    }
}
