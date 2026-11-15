using System;
using System.Collections.Generic;
using VFormat.Core.SpellCheck.Engine;
using VFormat.Core.SpellCheck.Models;
using VFormat.Core.SpellCheck.Processing;

namespace VFormat.Core.SpellCheck.Rules;

public class RepetitionChecker
{
	private readonly SpellDictionaryManager _dict;

	private static readonly HashSet<string> FunctionWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
	{
		"và", "của", "là", "các", "những", "được", "trong", "cho", "để", "với",
		"thì", "mà", "này", "đã", "sẽ", "đang", "bị", "theo", "về", "tại",
		"khi", "nếu", "nhưng", "hoặc", "vì", "do", "nên", "rằng", "cũng", "đều",
		"vào", "ở", "bởi", "tới", "đến", "hay", "như", "một", "từng", "qua",
		"trên", "dưới", "sau", "trước", "giữa", "đối", "với", "số", "ngày", "tháng",
		"năm", "điều", "khoản", "điểm", "chương", "mục", "phần"
	};

	private static readonly HashSet<string> AllowedDoubles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
	{
		"ai", "người", "nhà", "ngày", "tháng", "năm", "đời", "mãi", "đâu", "nơi",
		"lúc", "khắp", "đi", "nào", "thôi", "vâng", "dạ", "ừ", "ồ", "à",
		"ha", "hi", "hu", "chiều", "sáng", "tối", "đêm", "từ", "vừa", "rất",
		"lắm", "quá", "cứ", "luôn", "dần", "đều"
	};

	private static readonly HashSet<string> DateWords = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "ngày", "tháng", "năm" };

	public RepetitionChecker(SpellDictionaryManager dict)
	{
		_dict = dict;
	}

	public List<SpellCheckIssue> Check(List<TextToken> tokens, string normText, HashSet<int> skipTokenIndices)
	{
		List<SpellCheckIssue> list = new List<SpellCheckIssue>();
		if (tokens == null || tokens.Count < 2 || string.IsNullOrEmpty(normText))
		{
			return list;
		}
		for (int i = 0; i < tokens.Count - 1; i++)
		{
			TextToken textToken = tokens[i];
			if (textToken.Type != TokenType.Word || !textToken.Text.Equals("ngày", StringComparison.OrdinalIgnoreCase) || (skipTokenIndices != null && skipTokenIndices.Contains(i)))
			{
				continue;
			}
			int j;
			for (j = i + 1; j < tokens.Count && tokens[j].Type == TokenType.Whitespace; j++)
			{
			}
			if (j >= tokens.Count || tokens[j].Type != TokenType.Protected)
			{
				continue;
			}
			TextToken textToken2 = tokens[j];
			if (textToken2.Text.Length > 5 && textToken2.Text.StartsWith("ngày", StringComparison.OrdinalIgnoreCase) && char.IsWhiteSpace(textToken2.Text[4]) && WhitespaceOnly(normText, textToken.End, textToken2.Start))
			{
				int k;
				for (k = 4; k < textToken2.Text.Length && char.IsWhiteSpace(textToken2.Text[k]); k++)
				{
				}
				if (k < textToken2.Text.Length && char.IsDigit(textToken2.Text[k]))
				{
					int num = textToken2.Start + 4;
					list.Add(MakeIssue(textToken.Start, num, normText.Substring(textToken.Start, num - textToken.Start), textToken.Text, SpellCheckSeverity.Warning, "REPEAT_WORD", "Từ 'ngày' bị lặp trước ngày tháng, có thể do gõ thừa."));
				}
			}
		}
		List<int> list2 = new List<int>();
		for (int l = 0; l < tokens.Count; l++)
		{
			if (tokens[l].Type == TokenType.Word)
			{
				list2.Add(l);
			}
		}
		int num2 = 0;
		while (num2 < list2.Count - 1)
		{
			if (num2 + 3 < list2.Count && Adjacent(tokens, list2, num2, num2 + 3, normText) && SameWord(tokens[list2[num2]], tokens[list2[num2 + 2]]) && SameWord(tokens[list2[num2 + 1]], tokens[list2[num2 + 3]]) && !SameWord(tokens[list2[num2]], tokens[list2[num2 + 1]]) && !Skipped(skipTokenIndices, list2, num2, num2 + 3))
			{
				string text = tokens[list2[num2]].Text + " " + tokens[list2[num2 + 1]].Text;
				string text2 = text + " " + tokens[list2[num2 + 2]].Text + " " + tokens[list2[num2 + 3]].Text;
				if (!_dict.ContainsWord(text2.ToLowerInvariant()))
				{
					int start = tokens[list2[num2]].Start;
					int end = tokens[list2[num2 + 3]].End;
					list.Add(MakeIssue(start, end, normText.Substring(start, end - start), text, SpellCheckSeverity.Warning, "REPEAT_PHRASE", $"Cụm từ '{text}' bị lặp lại liền nhau, có thể do gõ thừa."));
					num2 += 4;
					continue;
				}
			}
			if (Adjacent(tokens, list2, num2, num2 + 1, normText) && SameWord(tokens[list2[num2]], tokens[list2[num2 + 1]]) && !Skipped(skipTokenIndices, list2, num2, num2 + 1))
			{
				TextToken textToken3 = tokens[list2[num2]];
				TextToken textToken4 = tokens[list2[num2 + 1]];
				string text3 = textToken3.Text.ToLowerInvariant();
				string word = text3 + " " + text3;
				int m;
				for (m = list2[num2 + 1] + 1; m < tokens.Count && tokens[m].Type == TokenType.Whitespace; m++)
				{
				}
				bool flag = m < tokens.Count && tokens[m].Type == TokenType.Number && DateWords.Contains(text3);
				if (flag && text3 == "năm" && num2 > 0)
				{
					switch (tokens[list2[num2 - 1]].Text.ToLowerInvariant())
					{
					case "trong":
					case "suốt":
					case "qua":
					case "sau":
					case "trước":
					case "khoảng":
					case "hơn":
					case "gần":
						flag = false;
						break;
					}
				}
				if ((flag || (!_dict.ContainsWord(word) && !AllowedDoubles.Contains(text3))) && !IsNumberLike(textToken3.Text))
				{
					bool flag2 = FunctionWords.Contains(text3) || flag;
					list.Add(MakeIssue(textToken3.Start, textToken4.End, normText.Substring(textToken3.Start, textToken4.End - textToken3.Start), textToken3.Text, flag2 ? SpellCheckSeverity.Warning : SpellCheckSeverity.Suggestion, flag2 ? "REPEAT_WORD" : "REPEAT_WORD_REVIEW", flag2 ? $"Từ '{textToken3.Text}' bị lặp hai lần liền nhau, có thể do gõ thừa." : $"Từ '{textToken3.Text}' đứng lặp hai lần. Nếu không phải dụng ý (láy, nhấn mạnh) thì nên bỏ bớt một lần."));
					num2 += 2;
					continue;
				}
			}
			num2++;
		}
		return list;
	}

	private static bool Adjacent(List<TextToken> tokens, List<int> w, int from, int to, string text)
	{
		for (int i = from; i < to; i++)
		{
			int end = tokens[w[i]].End;
			int start = tokens[w[i + 1]].Start;
			if (!WhitespaceOnly(text, end, start))
			{
				return false;
			}
		}
		return true;
	}

	private static bool WhitespaceOnly(string text, int start, int end)
	{
		if (start < 0 || end <= start || end > text.Length)
		{
			return false;
		}
		for (int i = start; i < end; i++)
		{
			if (text[i] != ' ' && text[i] != '\u00a0' && text[i] != '\t')
			{
				return false;
			}
		}
		return true;
	}

	private static bool SameWord(TextToken a, TextToken b)
	{
		return string.Equals(a.Text, b.Text, StringComparison.OrdinalIgnoreCase);
	}

	private static bool Skipped(HashSet<int> skip, List<int> w, int from, int to)
	{
		if (skip == null || skip.Count == 0)
		{
			return false;
		}
		for (int i = from; i <= to; i++)
		{
			if (skip.Contains(w[i]))
			{
				return true;
			}
		}
		return false;
	}

	private static bool IsNumberLike(string s)
	{
		for (int i = 0; i < s.Length; i++)
		{
			if (char.IsLetter(s[i]))
			{
				return false;
			}
		}
		return true;
	}

	private static SpellCheckIssue MakeIssue(int start, int end, string original, string suggestion, SpellCheckSeverity severity, string ruleId, string description)
	{
		SpellCheckIssue spellCheckIssue = new SpellCheckIssue();
		spellCheckIssue.Start = start;
		spellCheckIssue.End = end;
		spellCheckIssue.Original = original;
		spellCheckIssue.Type = SpellCheckIssueType.Repetition;
		spellCheckIssue.Severity = severity;
		spellCheckIssue.RuleId = ruleId;
		spellCheckIssue.Source = "Quy tắc phát hiện từ lặp của OfficeFormat";
		spellCheckIssue.Description = description;
		spellCheckIssue.Suggestions.Add(new SpellSuggestion(suggestion, 0.95, "Bỏ phần lặp"));
		return spellCheckIssue;
	}
}
