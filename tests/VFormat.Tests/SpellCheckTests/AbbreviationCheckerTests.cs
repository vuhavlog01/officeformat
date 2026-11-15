using System.Collections.Generic;
using Xunit;
using VFormat.Core.SpellCheck.Engine;
using VFormat.Core.SpellCheck.Processing;
using VFormat.Core.SpellCheck.Rules;
using VFormat.Core.SpellCheck.Models;

namespace VFormat.Tests.SpellCheckTests;

public class AbbreviationCheckerTests
{
    [Fact]
    public void AbbreviationChecker_SessionReset_ClearsSeenCache()
    {
        var dict = SpellDictionaryManager.Instance;
        var checker = new AbbreviationChecker(dict);

        // First session check
        var tokenizer = new VietnameseTokenizer();
        string text = "Quyết định của UBND tỉnh.";
        var tokens = tokenizer.Tokenize(text);

        var issues = checker.Check(tokens, text, SpellCheckProfile.LegalNormative);
        Assert.NotNull(issues);

        // Reset document session
        checker.ResetDocumentSession();

        // Second check after reset
        var issuesAfterReset = checker.Check(tokens, text, SpellCheckProfile.LegalNormative);
        Assert.NotNull(issuesAfterReset);
    }
}
