using System.Collections.Generic;
using Xunit;
using VFormat.Core.SpellCheck.Engine;
using VFormat.Core.SpellCheck.Processing;
using VFormat.Core.SpellCheck.Rules;
using VFormat.Core.SpellCheck.Models;

namespace VFormat.Tests.SpellCheckTests;

public class RepetitionCheckerTests
{
    [Fact]
    public void RepetitionChecker_DuplicateWord_DetectsRepetition()
    {
        var dict = SpellDictionaryManager.Instance;
        var checker = new RepetitionChecker(dict);
        var tokenizer = new VietnameseTokenizer();

        string text = "Căn cứ vào quyết định của của Ủy ban nhân dân.";
        List<TextToken> tokens = tokenizer.Tokenize(text);

        List<SpellCheckIssue> issues = checker.Check(tokens, text, null);

        Assert.NotEmpty(issues);
        Assert.Contains(issues, i => i.Type == SpellCheckIssueType.Repetition && i.Original.Trim() == "của của");
    }

    [Fact]
    public void RepetitionChecker_AllowedDouble_DoesNotFlag()
    {
        var dict = SpellDictionaryManager.Instance;
        var checker = new RepetitionChecker(dict);
        var tokenizer = new VietnameseTokenizer();

        // "nhà nhà" and "người người" are in AllowedDoubles
        string text = "Mọi người người đều hân hoan.";
        List<TextToken> tokens = tokenizer.Tokenize(text);

        List<SpellCheckIssue> issues = checker.Check(tokens, text, null);

        Assert.DoesNotContain(issues, i => i.Original.Trim() == "người người");
    }
}
