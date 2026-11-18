using System;
using System.Collections.Generic;
using System.Text;
using VFormat.Core.SpellCheck.Engine;
using VFormat.Core.SpellCheck.Models;
using VFormat.Core.SpellCheck.Processing;

namespace VFormat.Core.SpellCheck.Rules;

public class AbbreviationChecker
{
	private readonly SpellDictionaryManager _dictManager;

	private readonly HashSet<string> _seenAbbreviations;

	private static readonly Dictionary<string, string> AbbreviationFullForms = new Dictionary<string, string>(StringComparer.Ordinal)
	{
		{ "UBND", "Ủy ban nhân dân" },
		{ "HĐND", "Hội đồng nhân dân" },
		{ "CP", "Chính phủ" },
		{ "TTg", "Thủ tướng Chính phủ" },
		{ "BGDĐT", "Bộ Giáo dục và Đào tạo" },
		{ "BQP", "Bộ Quốc phòng" },
		{ "BNV", "Bộ Nội vụ" },
		{ "BTP", "Bộ Tư pháp" },
		{ "QH", "Quốc hội" },
		{ "UBTVQH", "Ủy ban Thường vụ Quốc hội" },
		{ "VPCP", "Văn phòng Chính phủ" },
		{ "BHXH", "Bảo hiểm xã hội" }
	};

	public AbbreviationChecker(SpellDictionaryManager dictManager)
	{
		_dictManager = dictManager;
		_seenAbbreviations = new HashSet<string>(StringComparer.Ordinal);
	}

	public void ResetDocumentSession()
	{
		_seenAbbreviations.Clear();
	}

	public List<SpellCheckIssue> Check(List<TextToken> tokens, string fullText, SpellCheckProfile profile)
	{
		List<SpellCheckIssue> list = new List<SpellCheckIssue>();
		if (tokens == null || tokens.Count == 0)
		{
			return list;
		}
		if (profile != SpellCheckProfile.LegalNormative)
		{
			return list;
		}
		for (int i = 0; i < tokens.Count; i++)
		{
			TextToken textToken = tokens[i];
			if (textToken.Type == TokenType.Protected || textToken.Type == TokenType.Whitespace || textToken.Type == TokenType.Punctuation)
			{
				continue;
			}
			string text = textToken.Text;
			if ((!_dictManager.IsAbbreviation(text) && !AbbreviationFullForms.ContainsKey(text)) || _seenAbbreviations.Contains(text))
			{
				continue;
			}
			AbbreviationFullForms.TryGetValue(text, out var value);
			if (!HasFullFormNearby(tokens, i, value))
			{
				SpellCheckIssue spellCheckIssue = new SpellCheckIssue();
				spellCheckIssue.Start = textToken.Start;
				spellCheckIssue.End = textToken.End;
				spellCheckIssue.Original = text;
				spellCheckIssue.Type = SpellCheckIssueType.Abbreviation;
				spellCheckIssue.Severity = SpellCheckSeverity.Warning;
				spellCheckIssue.RuleId = "LEGAL_ABBR_FIRST_USE";
				spellCheckIssue.Source = "Nghị định 78/2025/NĐ-CP, Điều 60, Khoản 4";
				spellCheckIssue.Description = $"Từ ngữ viết tắt '{text}' cần có dạng đầy đủ ở lần xuất hiện đầu tiên trong văn bản.";
				if (!string.IsNullOrEmpty(value))
				{
					spellCheckIssue.Suggestions.Add(new SpellSuggestion($"{value} ({text})", 0.9, "Bổ sung dạng đầy đủ lần đầu"));
				}
				list.Add(spellCheckIssue);
			}
			_seenAbbreviations.Add(text);
		}
		return list;
	}

	private static bool HasFullFormNearby(List<TextToken> tokens, int currentIndex, string fullForm)
	{
		if (string.IsNullOrEmpty(fullForm))
		{
			return true;
		}
		int num = Math.Max(0, currentIndex - 15);
		StringBuilder stringBuilder = new StringBuilder();
		for (int i = num; i < currentIndex; i++)
		{
			stringBuilder.Append(tokens[i].Text);
		}
		string text = stringBuilder.ToString();
		if (text.IndexOf(fullForm, StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return true;
		}
		int num2 = Math.Min(tokens.Count, currentIndex + 15);
		stringBuilder.Length = 0;
		for (int j = currentIndex + 1; j < num2; j++)
		{
			stringBuilder.Append(tokens[j].Text);
		}
		string text2 = stringBuilder.ToString();
		if (text2.IndexOf(fullForm, StringComparison.OrdinalIgnoreCase) >= 0)
		{
			return true;
		}
		return false;
	}
}
