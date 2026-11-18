using System;
using System.Collections.Generic;
using System.Text;
using VFormat.Core.SpellCheck.Engine;
using VFormat.Core.SpellCheck.Models;
using VFormat.Core.SpellCheck.Processing;

namespace VFormat.Core.SpellCheck.Rules;

public class ContextualConfusionChecker
{
	private readonly SpellDictionaryManager _dictManager;

	public ContextualConfusionChecker(SpellDictionaryManager dictManager)
	{
		_dictManager = dictManager;
	}

	public List<SpellCheckIssue> Check(List<TextToken> tokens, string normText, SpellCheckProfile profile, out HashSet<int> matchedTokenIndices)
	{
		List<SpellCheckIssue> list = new List<SpellCheckIssue>();
		matchedTokenIndices = new HashSet<int>();
		if (tokens == null || tokens.Count == 0)
		{
			return list;
		}
		int count = tokens.Count;
		int num = 0;
		while (num < count)
		{
			TextToken textToken = tokens[num];
			if (textToken.Type != TokenType.Word)
			{
				num++;
				continue;
			}
			List<int> list2 = new List<int>();
			list2.Add(num);
			int currentIndex = num;
			for (int i = 1; i < _dictManager.MaxConfusablePhraseWords; i++)
			{
				int nextWordIndex = GetNextWordIndex(tokens, currentIndex);
				if (nextWordIndex == -1)
				{
					break;
				}
				list2.Add(nextWordIndex);
				currentIndex = nextWordIndex;
			}
			bool flag = false;
			for (int num2 = list2.Count; num2 >= 2; num2--)
			{
				StringBuilder stringBuilder = new StringBuilder();
				for (int j = 0; j < num2; j++)
				{
					if (j > 0)
					{
						stringBuilder.Append(' ');
					}
					stringBuilder.Append(tokens[list2[j]].Text);
				}
				string text = stringBuilder.ToString();
				if (_dictManager.TryGetConfusablePhrase(text, out var entry) && !_dictManager.IsUserAccepted(text) && ContextAllows(tokens, list2[0], list2[num2 - 1], entry))
				{
					TextToken startToken = tokens[list2[0]];
					TextToken endToken = tokens[list2[num2 - 1]];
					SpellCheckIssue item = CreateIssue(startToken, endToken, normText, text, entry);
					list.Add(item);
					for (int k = 0; k < num2; k++)
					{
						matchedTokenIndices.Add(list2[k]);
					}
					num = list2[num2 - 1] + 1;
					flag = true;
					break;
				}
			}
			if (!flag)
			{
				num++;
			}
		}
		return list;
	}

	private static bool ContextAllows(List<TextToken> tokens, int firstIndex, int lastIndex, ConfusablePhraseEntry entry)
	{
		if (entry.not_after != null)
		{
			int prevWordIndex = GetPrevWordIndex(tokens, firstIndex);
			if (prevWordIndex != -1 && entry.not_after.Contains(NormalizeWord(tokens[prevWordIndex].Text)))
			{
				return false;
			}
		}
		int nextWordIndex = GetNextWordIndex(tokens, lastIndex);
		string text = ((nextWordIndex != -1) ? NormalizeWord(tokens[nextWordIndex].Text) : null);
		if (entry.not_before != null && text != null && entry.not_before.Contains(text))
		{
			return false;
		}
		if (entry.only_before != null)
		{
			if (text == null)
			{
				return entry.only_before.Contains("<END>");
			}
			return entry.only_before.Contains(text);
		}
		return true;
	}

	private static string NormalizeWord(string text)
	{
		return VietnameseUnicodeNormalizer.NormalizeNfc(text).ToLowerInvariant();
	}

	private static int GetPrevWordIndex(List<TextToken> tokens, int currentIndex)
	{
		for (int num = currentIndex - 1; num >= 0; num--)
		{
			if (tokens[num].Type != TokenType.Whitespace)
			{
				if (tokens[num].Type != TokenType.Word)
				{
					break;
				}
				return num;
			}
		}
		return -1;
	}

	private static int GetNextWordIndex(List<TextToken> tokens, int currentIndex)
	{
		for (int i = currentIndex + 1; i < tokens.Count; i++)
		{
			if (tokens[i].Type != TokenType.Whitespace)
			{
				if (tokens[i].Type != TokenType.Word)
				{
					break;
				}
				return i;
			}
		}
		return -1;
	}

	private static SpellCheckIssue CreateIssue(TextToken startToken, TextToken endToken, string fullText, string phrase, ConfusablePhraseEntry entry)
	{
		int start = startToken.Start;
		int end = endToken.End;
		string text = fullText.Substring(start, end - start);
		SpellCheckSeverity severity = SpellCheckSeverity.Error;
		if (!string.IsNullOrEmpty(entry.severity) && entry.severity.Equals("warning", StringComparison.OrdinalIgnoreCase))
		{
			severity = SpellCheckSeverity.Warning;
		}
		SpellCheckIssue spellCheckIssue = new SpellCheckIssue();
		spellCheckIssue.Start = start;
		spellCheckIssue.End = end;
		spellCheckIssue.Original = text;
		spellCheckIssue.Type = SpellCheckIssueType.Spelling;
		spellCheckIssue.Severity = severity;
		spellCheckIssue.RuleId = "CONTEXTUAL_CONFUSION";
		spellCheckIssue.Source = ((!string.IsNullOrEmpty(entry.source)) ? entry.source : "Từ điển tiếng Việt Viện Ngôn ngữ học");
		spellCheckIssue.Description = ((!string.IsNullOrEmpty(entry.description)) ? entry.description : $"Cụm từ dễ nhầm lẫn: nên dùng '{entry.replacement}' thay cho '{text}'.");
		spellCheckIssue.WhenToUse = entry.when_to_use;
		spellCheckIssue.WhenNotToUse = entry.when_not_to_use;
		spellCheckIssue.UsageGuide = entry.usage_guide;
		if (entry.examples != null && entry.examples.Count > 0)
		{
			spellCheckIssue.Examples = new List<string>(entry.examples);
		}
		string text2 = MatchCase(text, entry.replacement);
		spellCheckIssue.Suggestions.Add(new SpellSuggestion(text2, 0.98, "Từ ngữ chuẩn mực trong văn bản"));
		return spellCheckIssue;
	}

	private static string MatchCase(string original, string target)
	{
		if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(target))
		{
			return target;
		}
		bool flag = true;
		foreach (char c in original)
		{
			if (char.IsLetter(c) && !char.IsUpper(c) && c != 'Đ')
			{
				flag = false;
				break;
			}
		}
		if (flag && original.Length > 1)
		{
			return target.ToUpperInvariant();
		}
		if (char.IsUpper(original[0]))
		{
			return char.ToUpperInvariant(target[0]) + ((target.Length > 1) ? target.Substring(1) : string.Empty);
		}
		return target;
	}
}
