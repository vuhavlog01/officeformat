using System;
using System.Collections.Generic;
using VFormat.Core.SpellCheck.Models;
using VFormat.Core.SpellCheck.Processing;

namespace VFormat.Core.SpellCheck.Rules;

public class VariantChecker
{
	private static readonly Dictionary<string, string> EducationPreferences = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
	{
		{ "kỹ", "kĩ" },
		{ "hy", "hi" },
		{ "kỷ", "kỉ" },
		{ "lý", "lí" },
		{ "mỹ", "mĩ" },
		{ "tỷ", "tỉ" },
		{ "sỹ", "sĩ" }
	};

	public List<SpellCheckIssue> Check(List<TextToken> tokens, SpellCheckProfile profile)
	{
		List<SpellCheckIssue> list = new List<SpellCheckIssue>();
		if (tokens == null || tokens.Count == 0)
		{
			return list;
		}
		switch (profile)
		{
		case SpellCheckProfile.Administrative:
		case SpellCheckProfile.Military:
		case SpellCheckProfile.Party:
			return list;
		case SpellCheckProfile.EducationOrthography:
		{
			for (int i = 0; i < tokens.Count; i++)
			{
				TextToken textToken = tokens[i];
				if (textToken.Type == TokenType.Word)
				{
					string key = textToken.Text.ToLowerInvariant();
					if (EducationPreferences.TryGetValue(key, out var value))
					{
						string text = (char.IsUpper(textToken.Text[0]) ? (char.ToUpperInvariant(value[0]) + ((value.Length > 1) ? value.Substring(1) : string.Empty)) : value);
						SpellCheckIssue spellCheckIssue = new SpellCheckIssue();
						spellCheckIssue.Start = textToken.Start;
						spellCheckIssue.End = textToken.End;
						spellCheckIssue.Original = textToken.Text;
						spellCheckIssue.Type = SpellCheckIssueType.Variant;
						spellCheckIssue.Severity = SpellCheckSeverity.Suggestion;
						spellCheckIssue.RuleId = "EDU_ORTHOGRAPHY_I_Y";
						spellCheckIssue.Source = "Quyết định 1989/QĐ-BGDĐT, Điều 9";
						spellCheckIssue.Description = $"Quy định chính tả ngành giáo dục ưu tiên viết '{text}' thay cho '{textToken.Text}'.";
						spellCheckIssue.Suggestions.Add(new SpellSuggestion(text, 0.9, "Biến thể ưu tiên theo QĐ 1989"));
						list.Add(spellCheckIssue);
					}
				}
			}
			break;
		}
		}
		return list;
	}
}
