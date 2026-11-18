using System;
using System.Collections.Generic;
using VFormat.Core.SpellCheck.Models;
using VFormat.Core.SpellCheck.Processing;

namespace VFormat.Core.SpellCheck.Rules;

public class CapitalizationChecker
{
	private static readonly HashSet<string> CompoundHeads = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
	{
		"danh", "hạng", "đề", "tiết", "chuyên", "tiểu", "thành", "bộ", "cổ", "từng",
		"phân", "khẩu"
	};

	private static readonly HashSet<string> PeriodAbbreviations = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
	{
		"TP", "Tp", "ThS", "TS", "PGS", "GS", "BS", "KS", "CN", "Tr",
		"St", "vd", "VD", "Th", "Q", "P", "Ths", "Đ/c", "đ/c"
	};

	private static readonly HashSet<string> ConnectingConjunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "và", "hoặc", "hay", "nhưng", "song", "mà", "cũng", "với" };

	public List<SpellCheckIssue> Check(List<TextToken> tokens, string text, SpellCheckProfile profile)
	{
		List<SpellCheckIssue> list = new List<SpellCheckIssue>();
		if (tokens == null || tokens.Count == 0)
		{
			return list;
		}
		bool flag = true;
		for (int i = 0; i < tokens.Count; i++)
		{
			TextToken textToken = tokens[i];
			if (textToken.Type == TokenType.Whitespace)
			{
				continue;
			}
			if (textToken.Type == TokenType.Punctuation)
			{
				if (textToken.Text == "\n" || IsSentenceTerminator(text, tokens, i))
				{
					flag = true;
				}
			}
			else if (textToken.Type == TokenType.Protected || textToken.Type == TokenType.Number)
			{
				flag = false;
			}
			else
			{
				if (textToken.Type != TokenType.Word)
				{
					continue;
				}
				string text2 = textToken.Text;
				if (flag && char.IsLower(text2[0]))
				{
					if (text2.Length == 1 && i + 1 < tokens.Count)
					{
						TextToken nextNonWhitespaceToken = GetNextNonWhitespaceToken(tokens, i);
						if (nextNonWhitespaceToken != null && (nextNonWhitespaceToken.Text == "." || nextNonWhitespaceToken.Text == ")" || nextNonWhitespaceToken.Text == "-" || nextNonWhitespaceToken.Text == "/"))
						{
							flag = true;
							continue;
						}
					}
					if (IsConnectingConjunction(text2))
					{
						flag = false;
						continue;
					}
					string text3 = char.ToUpperInvariant(text2[0]) + ((text2.Length > 1) ? text2.Substring(1) : string.Empty);
					SpellCheckIssue spellCheckIssue = new SpellCheckIssue();
					spellCheckIssue.Start = textToken.Start;
					spellCheckIssue.End = textToken.End;
					spellCheckIssue.Original = text2;
					spellCheckIssue.Type = SpellCheckIssueType.Capitalization;
					spellCheckIssue.Severity = SpellCheckSeverity.Warning;
					spellCheckIssue.RuleId = "CAP_SENTENCE_START";
					spellCheckIssue.Source = "Nghị định 30/2020/NĐ-CP, Phụ lục II, Mục I";
					spellCheckIssue.Description = "Chữ cái đầu câu phải viết hoa.";
					spellCheckIssue.Suggestions.Add(new SpellSuggestion(text3, 0.95, "Viết hoa chữ cái đầu câu"));
					list.Add(spellCheckIssue);
				}
				if (char.IsLower(text2[0]))
				{
					switch (text2.ToLowerInvariant())
					{
					case "điều":
					case "chương":
					case "mục":
					case "phần":
						if (i + 1 < tokens.Count && !IsCompoundHead(GetPrevWord(tokens, i)))
						{
							TextToken nextNonWhitespaceToken2 = GetNextNonWhitespaceToken(tokens, i);
							if (nextNonWhitespaceToken2 != null && (nextNonWhitespaceToken2.Type == TokenType.Number || (nextNonWhitespaceToken2.Type == TokenType.Word && (IsRomanNumeral(nextNonWhitespaceToken2.Text) || char.IsDigit(nextNonWhitespaceToken2.Text[0])))))
							{
								string text4 = char.ToUpperInvariant(text2[0]) + text2.Substring(1);
								SpellCheckIssue spellCheckIssue2 = new SpellCheckIssue();
								spellCheckIssue2.Start = textToken.Start;
								spellCheckIssue2.End = textToken.End;
								spellCheckIssue2.Original = text2;
								spellCheckIssue2.Type = SpellCheckIssueType.Capitalization;
								spellCheckIssue2.Severity = SpellCheckSeverity.Warning;
								spellCheckIssue2.RuleId = "CAP_LEGAL_STRUCT_REF";
								spellCheckIssue2.Source = "Nghị định 30/2020/NĐ-CP, Phụ lục II, Mục V.7";
								spellCheckIssue2.Description = $"Viết hoa chữ cái đầu của {text4} khi viện dẫn điều khoản văn bản cụ thể.";
								spellCheckIssue2.Suggestions.Add(new SpellSuggestion(text4, 0.95, "Viết hoa thành tố viện dẫn"));
								list.Add(spellCheckIssue2);
							}
						}
						break;
					}
				}
				flag = false;
			}
		}
		return list;
	}

	private static bool IsCompoundHead(string previousWord)
	{
		if (previousWord != null)
		{
			return CompoundHeads.Contains(previousWord);
		}
		return false;
	}

	private static string GetPrevWord(List<TextToken> tokens, int currentIndex)
	{
		for (int num = currentIndex - 1; num >= 0; num--)
		{
			if (tokens[num].Type != TokenType.Whitespace)
			{
				if (tokens[num].Type != TokenType.Word)
				{
					return null;
				}
				return tokens[num].Text;
			}
		}
		return null;
	}

	private static bool IsSentenceTerminator(string text, List<TextToken> tokens, int index)
	{
		TextToken textToken = tokens[index];
		string text2 = textToken.Text;
		if (text2 != "." && text2 != "!" && text2 != "?")
		{
			return false;
		}
		if (string.IsNullOrEmpty(text))
		{
			return true;
		}
		int start = textToken.Start;
		int end = textToken.End;
		if (text2 == ".")
		{
			char c = ((start > 0) ? text[start - 1] : ' ');
			char c2 = ((end < text.Length) ? text[end] : ' ');
			if (c == '.' || c == '…' || c2 == '.' || c2 == '…')
			{
				return false;
			}
			string text3 = ((index > 0 && tokens[index - 1].Type == TokenType.Word) ? tokens[index - 1].Text : null);
			if (text3 != null)
			{
				if (PeriodAbbreviations.Contains(text3))
				{
					return false;
				}
				if (text3.Length == 1 && tokens[index - 1].Start > 0 && text[tokens[index - 1].Start - 1] == '.')
				{
					return false;
				}
			}
		}
		int i;
		for (i = end; i < text.Length && IsClosingMark(text[i]); i++)
		{
		}
		if (i < text.Length)
		{
			return char.IsWhiteSpace(text[i]);
		}
		return true;
	}

	private static bool IsClosingMark(char c)
	{
		if (c != '"' && c != '\'' && c != '”' && c != '’' && c != ')' && c != ']')
		{
			return c == '»';
		}
		return true;
	}

	private static TextToken GetNextNonWhitespaceToken(List<TextToken> tokens, int currentIndex)
	{
		for (int i = currentIndex + 1; i < tokens.Count; i++)
		{
			if (tokens[i].Type != TokenType.Whitespace)
			{
				return tokens[i];
			}
		}
		return null;
	}

	private static bool IsConnectingConjunction(string word)
	{
		if (string.IsNullOrEmpty(word))
		{
			return false;
		}
		return ConnectingConjunctions.Contains(word);
	}

	private static bool IsRomanNumeral(string s)
	{
		if (string.IsNullOrEmpty(s))
		{
			return false;
		}
		string text = s.ToUpperInvariant();
		foreach (char c in text)
		{
			if (c != 'I' && c != 'V' && c != 'X' && c != 'L' && c != 'C' && c != 'D' && c != 'M')
			{
				return false;
			}
		}
		return true;
	}
}
