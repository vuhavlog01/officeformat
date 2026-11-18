using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading;
using VFormat.Core.SpellCheck.Engine;
using VFormat.Core.SpellCheck.Models;
using VFormat.Core.SpellCheck.Processing;
using VFormat.Core.SpellCheck.Rules;
using VFormat.Core.SpellCheck.Validation;

namespace VFormat.Core.SpellCheck;

public class SpellCheckService
{
	private static SpellCheckService _instance;

	private static readonly object _lock = new object();

	private readonly SpellDictionaryManager _dictManager;

	private readonly ProtectedPatternMatcher _protectedMatcher;

	private readonly VietnameseTokenizer _tokenizer;

	private readonly CandidateGenerator _candidateGenerator;

	private readonly CandidateRanker _candidateRanker;

	private readonly CapitalizationChecker _capitalizationChecker;

	private readonly AbbreviationChecker _abbreviationChecker;

	private readonly VariantChecker _variantChecker;

	private readonly ContextualConfusionChecker _contextualChecker;

	private readonly RepetitionChecker _repetitionChecker;

	private static readonly HashSet<char> ViVowels = new HashSet<char>("aàáảãạăằắẳẵặâầấẩẫậeèéẻẽẹêềếểễệiìíỉĩịoòóỏõọôồốổỗộơờớởỡợuùúủũụưừứửữựyỳýỷỹỵAÀÁẢÃẠĂẰẮẲẴẶÂẦẤẨẪẬEÈÉẺẼẸÊỀẾỂỄỆIÌÍỈĨỊOÒÓỎÕỌÔỒỐỔỖỘƠỜỚỞỠỢUÙÚỦŨỤƯỪỨỬỮỰYỲÝỶỸỴ");

	private static readonly Regex RomanNumeralRegex = new Regex("^M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})$", RegexOptions.IgnoreCase);

	public static SpellCheckService Instance
	{
		get
		{
			if (_instance == null)
			{
				lock (_lock)
				{
					if (_instance == null)
					{
						_instance = new SpellCheckService();
					}
				}
			}
			return _instance;
		}
	}

	public SpellDictionaryManager DictionaryManager => _dictManager;

	public SpellCheckService(string customDataDir = null)
	{
		_dictManager = new SpellDictionaryManager(customDataDir);
		_protectedMatcher = new ProtectedPatternMatcher();
		_tokenizer = new VietnameseTokenizer(_protectedMatcher);
		_candidateGenerator = new CandidateGenerator(_dictManager);
		_candidateRanker = new CandidateRanker(_dictManager);
		_capitalizationChecker = new CapitalizationChecker();
		_abbreviationChecker = new AbbreviationChecker(_dictManager);
		_variantChecker = new VariantChecker();
		_contextualChecker = new ContextualConfusionChecker(_dictManager);
		_repetitionChecker = new RepetitionChecker(_dictManager);
	}

	public void ResetDocumentSession()
	{
		_abbreviationChecker.ResetDocumentSession();
	}

	public void AddCustomWord(string word)
	{
		_dictManager.AddUserWord(word);
	}

	public void IgnoreWord(string word)
	{
		_dictManager.IgnoreWord(word);
	}

	public List<SpellCheckIssue> CheckParagraphs(IList<string> paragraphs, SpellCheckProfile profile = SpellCheckProfile.Administrative)
	{
		ResetDocumentSession();
		List<SpellCheckIssue> list = new List<SpellCheckIssue>();
		if (paragraphs == null || paragraphs.Count == 0)
		{
			return list;
		}
		for (int i = 0; i < paragraphs.Count; i++)
		{
			string text = paragraphs[i];
			if (!string.IsNullOrEmpty(text))
			{
				List<SpellCheckIssue> collection = CheckParagraph(text, i, profile);
				list.AddRange(collection);
			}
		}
		return list;
	}

	public List<SpellCheckIssue> CheckText(string text, SpellCheckProfile profile = SpellCheckProfile.Administrative)
	{
		ResetDocumentSession();
		return CheckParagraph(text, 0, profile);
	}

	public List<SpellCheckIssue> CheckParagraph(string text, int paragraphIndex, SpellCheckProfile profile = SpellCheckProfile.Administrative)
	{
		return CheckParagraph(text, paragraphIndex, profile, CancellationToken.None);
	}

	public List<SpellCheckIssue> CheckParagraph(string text, int paragraphIndex, SpellCheckProfile profile, CancellationToken cancellationToken)
	{
		cancellationToken.ThrowIfCancellationRequested();
		List<SpellCheckIssue> list = new List<SpellCheckIssue>();
		if (string.IsNullOrEmpty(text))
		{
			return list;
		}
		NormalizedMapping normalizedMapping = VietnameseUnicodeNormalizer.NormalizeWithMapping(text);
		string normalizedText = normalizedMapping.NormalizedText;
		List<TextToken> list2 = _tokenizer.Tokenize(normalizedText);
		cancellationToken.ThrowIfCancellationRequested();
		List<SpellCheckIssue> list3 = _contextualChecker.Check(list2, normalizedText, profile, out var matchedTokenIndices);
		for (int i = 0; i < list3.Count; i++)
		{
			SpellCheckIssue spellCheckIssue = list3[i];
			spellCheckIssue.Start = normalizedMapping.ToOriginalStart(spellCheckIssue.Start);
			spellCheckIssue.End = normalizedMapping.ToOriginalEnd(spellCheckIssue.End);
			spellCheckIssue.ParagraphIndex = paragraphIndex;
			spellCheckIssue.ContextSnippet = ExtractContextSnippet(text, spellCheckIssue.Start, spellCheckIssue.End);
			list.Add(spellCheckIssue);
		}
		for (int j = 0; j < list2.Count; j++)
		{
			cancellationToken.ThrowIfCancellationRequested();
			if (matchedTokenIndices.Contains(j))
			{
				continue;
			}
			TextToken textToken = list2[j];
			if (textToken.Type != TokenType.Word)
			{
				continue;
			}
			string text2 = textToken.Text;
			string text3 = text2.ToLowerInvariant();
			if (_dictManager.ContainsWord(text2) || _dictManager.ContainsWord(text3) || _dictManager.IsAbbreviation(text2) || _dictManager.IsProperName(text2) || _dictManager.IsProperName(text3) || _dictManager.IsForeignTerm(text2) || _dictManager.IsForeignTerm(text3))
			{
				continue;
			}
			if (text2.Contains("-") || text2.Contains("'") || text2.Contains("’") || text2.Contains("‘") || text2.Contains("/"))
			{
				List<KeyValuePair<int, string>> list4 = SplitCompoundParts(text2);
				if (list4.Count >= 2)
				{
					List<KeyValuePair<int, string>> list5 = new List<KeyValuePair<int, string>>();
					for (int k = 0; k < list4.Count; k++)
					{
						if (!IsAcceptableCompoundPart(list4[k].Value))
						{
							list5.Add(list4[k]);
						}
					}
					if (list5.Count != 0)
					{
						string adjacentWord = GetAdjacentWord(list2, j, -1);
						string adjacentWord2 = GetAdjacentWord(list2, j, 1);
						for (int l = 0; l < list5.Count; l++)
						{
							int num = textToken.Start + list5[l].Key;
							int normEnd = num + list5[l].Value.Length;
							list.Add(CreateLexicalIssue(text, normalizedMapping, num, normEnd, list5[l].Value, adjacentWord, adjacentWord2, paragraphIndex, null, cancellationToken));
						}
					}
					continue;
				}
			}
			if (IsFileNameOrIdentifierPart(normalizedText, textToken) || (text2.Length == 1 && char.IsLetter(text2[0])) || IsRomanNumeral(text2))
			{
				continue;
			}
			string overrideDescription = null;
			if (IsAllUpper(text2) && text2.Length >= 2)
			{
				bool flag = IsCommonAsciiAcronym(text2);
				if (flag && IsLikelyUnaccentedVietnamese(text3))
				{
					overrideDescription = "Chữ in hoa không dấu: có thể thiếu dấu tiếng Việt.";
				}
				else if (!HasVietnameseVowel(text2) || flag)
				{
					continue;
				}
			}
			string adjacentWord3 = GetAdjacentWord(list2, j, -1);
			string adjacentWord4 = GetAdjacentWord(list2, j, 1);
			list.Add(CreateLexicalIssue(text, normalizedMapping, textToken.Start, textToken.End, text2, adjacentWord3, adjacentWord4, paragraphIndex, overrideDescription, cancellationToken));
		}
		cancellationToken.ThrowIfCancellationRequested();
		List<SpellCheckIssue> list6 = _capitalizationChecker.Check(list2, normalizedText, profile);
		for (int m = 0; m < list6.Count; m++)
		{
			SpellCheckIssue spellCheckIssue2 = list6[m];
			spellCheckIssue2.Start = normalizedMapping.ToOriginalStart(spellCheckIssue2.Start);
			spellCheckIssue2.End = normalizedMapping.ToOriginalEnd(spellCheckIssue2.End);
			spellCheckIssue2.ParagraphIndex = paragraphIndex;
			spellCheckIssue2.ContextSnippet = ExtractContextSnippet(text, spellCheckIssue2.Start, spellCheckIssue2.End);
			list.Add(spellCheckIssue2);
		}
		cancellationToken.ThrowIfCancellationRequested();
		List<SpellCheckIssue> list7 = _abbreviationChecker.Check(list2, normalizedText, profile);
		for (int n = 0; n < list7.Count; n++)
		{
			SpellCheckIssue spellCheckIssue3 = list7[n];
			spellCheckIssue3.Start = normalizedMapping.ToOriginalStart(spellCheckIssue3.Start);
			spellCheckIssue3.End = normalizedMapping.ToOriginalEnd(spellCheckIssue3.End);
			spellCheckIssue3.ParagraphIndex = paragraphIndex;
			spellCheckIssue3.ContextSnippet = ExtractContextSnippet(text, spellCheckIssue3.Start, spellCheckIssue3.End);
			list.Add(spellCheckIssue3);
		}
		cancellationToken.ThrowIfCancellationRequested();
		List<SpellCheckIssue> list8 = _variantChecker.Check(list2, profile);
		for (int num2 = 0; num2 < list8.Count; num2++)
		{
			SpellCheckIssue spellCheckIssue4 = list8[num2];
			spellCheckIssue4.Start = normalizedMapping.ToOriginalStart(spellCheckIssue4.Start);
			spellCheckIssue4.End = normalizedMapping.ToOriginalEnd(spellCheckIssue4.End);
			spellCheckIssue4.ParagraphIndex = paragraphIndex;
			spellCheckIssue4.ContextSnippet = ExtractContextSnippet(text, spellCheckIssue4.Start, spellCheckIssue4.End);
			list.Add(spellCheckIssue4);
		}
		cancellationToken.ThrowIfCancellationRequested();
		List<SpellCheckIssue> list9 = _repetitionChecker.Check(list2, normalizedText, matchedTokenIndices);
		for (int num3 = 0; num3 < list9.Count; num3++)
		{
			SpellCheckIssue spellCheckIssue5 = list9[num3];
			spellCheckIssue5.Start = normalizedMapping.ToOriginalStart(spellCheckIssue5.Start);
			spellCheckIssue5.End = normalizedMapping.ToOriginalEnd(spellCheckIssue5.End);
			spellCheckIssue5.Original = text.Substring(spellCheckIssue5.Start, spellCheckIssue5.End - spellCheckIssue5.Start);
			spellCheckIssue5.ParagraphIndex = paragraphIndex;
			spellCheckIssue5.ContextSnippet = ExtractContextSnippet(text, spellCheckIssue5.Start, spellCheckIssue5.End);
			list.Add(spellCheckIssue5);
		}
		list.RemoveAll((SpellCheckIssue issue) => _dictManager.IsIgnored(issue.Original));
		list.Sort((SpellCheckIssue a, SpellCheckIssue b) => a.Start.CompareTo(b.Start));
		return list;
	}

	private SpellCheckIssue CreateLexicalIssue(string originalText, NormalizedMapping mapping, int normStart, int normEnd, string rawWord, string prevWord, string nextWord, int paragraphIndex, string overrideDescription, CancellationToken cancellationToken)
	{
		string word = rawWord.ToLowerInvariant();
		List<string> candidates = _candidateGenerator.GenerateCandidates(rawWord, prevWord, nextWord, cancellationToken);
		List<SpellSuggestion> list = _candidateRanker.RankCandidates(rawWord, candidates, prevWord, nextWord);
		SyllableValidationResult syllableValidationResult = SyllableValidator.Validate(word);
		SpellCheckSeverity spellCheckSeverity = SpellCheckSeverity.Error;
		string description = "Từ không có trong từ điển tiếng Việt.";
		if (list.Count > 0 && list[0].Score >= 0.9)
		{
			spellCheckSeverity = SpellCheckSeverity.Error;
			description = $"Có thể bạn muốn viết: '{list[0].Text}'";
		}
		else if (syllableValidationResult == SyllableValidationResult.ValidSyllable)
		{
			spellCheckSeverity = SpellCheckSeverity.Warning;
			description = "Từ lạ hoặc chưa có trong từ điển cơ quan.";
		}
		else if (syllableValidationResult == SyllableValidationResult.ForeignOrUnknown || IsLikelyForeignWord(rawWord))
		{
			spellCheckSeverity = SpellCheckSeverity.Info;
			description = "Thuật ngữ nước ngoài hoặc ký hiệu chưa được định nghĩa.";
			if (string.IsNullOrEmpty(overrideDescription))
			{
				list = new List<SpellSuggestion>();
			}
		}
		if (!string.IsNullOrEmpty(overrideDescription))
		{
			if (spellCheckSeverity == SpellCheckSeverity.Error && (list.Count == 0 || list[0].Score < 0.9))
			{
				spellCheckSeverity = SpellCheckSeverity.Warning;
			}
			if (spellCheckSeverity == SpellCheckSeverity.Info)
			{
				spellCheckSeverity = SpellCheckSeverity.Warning;
			}
			description = overrideDescription;
		}
		SpellCheckIssue spellCheckIssue = new SpellCheckIssue();
		spellCheckIssue.Start = mapping.ToOriginalStart(normStart);
		spellCheckIssue.End = mapping.ToOriginalEnd(normEnd);
		spellCheckIssue.ParagraphIndex = paragraphIndex;
		spellCheckIssue.Original = originalText.Substring(spellCheckIssue.Start, spellCheckIssue.End - spellCheckIssue.Start);
		spellCheckIssue.Type = SpellCheckIssueType.Spelling;
		spellCheckIssue.Severity = spellCheckSeverity;
		spellCheckIssue.Suggestions = list;
		spellCheckIssue.Description = description;
		spellCheckIssue.ContextSnippet = ExtractContextSnippet(originalText, spellCheckIssue.Start, spellCheckIssue.End);
		return spellCheckIssue;
	}

	private static List<KeyValuePair<int, string>> SplitCompoundParts(string rawWord)
	{
		List<KeyValuePair<int, string>> list = new List<KeyValuePair<int, string>>();
		int num = -1;
		for (int i = 0; i <= rawWord.Length; i++)
		{
			if (i == rawWord.Length || rawWord[i] == '-' || rawWord[i] == '\'' || rawWord[i] == '’' || rawWord[i] == '‘' || rawWord[i] == '/')
			{
				if (num >= 0)
				{
					list.Add(new KeyValuePair<int, string>(num, rawWord.Substring(num, i - num)));
				}
				num = -1;
			}
			else if (num < 0)
			{
				num = i;
			}
		}
		return list;
	}

	private bool IsAcceptableCompoundPart(string part)
	{
		if (string.IsNullOrEmpty(part))
		{
			return true;
		}
		string word = part.ToLowerInvariant();
		if (_dictManager.ContainsWord(part) || _dictManager.ContainsWord(word))
		{
			return true;
		}
		if (_dictManager.IsProperName(part) || _dictManager.IsAbbreviation(part) || _dictManager.IsForeignTerm(part))
		{
			return true;
		}
		if (part.Length == 1 && char.IsLetter(part[0]))
		{
			return true;
		}
		if (part.Length >= 2 && IsAllUpper(part))
		{
			return true;
		}
		SyllableValidationResult syllableValidationResult = SyllableValidator.Validate(word);
		if (syllableValidationResult == SyllableValidationResult.ForeignOrUnknown)
		{
			return true;
		}
		if (char.IsUpper(part[0]) && syllableValidationResult == SyllableValidationResult.ValidSyllable)
		{
			return true;
		}
		return false;
	}

	private bool IsLikelyUnaccentedVietnamese(string lowerWord)
	{
		if (string.IsNullOrEmpty(lowerWord) || lowerWord.Length < 2)
		{
			return false;
		}
		if (_dictManager.ContainsWord(lowerWord))
		{
			return false;
		}
		if (SyllableValidator.Validate(lowerWord) != SyllableValidationResult.ValidSyllable)
		{
			return false;
		}
		return _dictManager.HasAccentedVariant(lowerWord);
	}

	private static bool IsFileNameOrIdentifierPart(string text, TextToken token)
	{
		if (string.IsNullOrEmpty(text))
		{
			return false;
		}
		char c = ((token.Start > 0) ? text[token.Start - 1] : ' ');
		char c2 = ((token.End < text.Length) ? text[token.End] : ' ');
		if (c == '_' || c2 == '_')
		{
			return true;
		}
		if (c == '.' && token.Text.Length <= 5)
		{
			for (int i = 0; i < token.Text.Length; i++)
			{
				char c3 = char.ToLowerInvariant(token.Text[i]);
				if (c3 < 'a' || c3 > 'z')
				{
					return false;
				}
			}
			return true;
		}
		return false;
	}

	private static bool IsLikelyForeignWord(string rawWord)
	{
		if (string.IsNullOrEmpty(rawWord) || rawWord.Length < 2)
		{
			return false;
		}
		for (int i = 0; i < rawWord.Length; i++)
		{
			char c = char.ToLowerInvariant(rawWord[i]);
			if (c < 'a' || c > 'z')
			{
				return false;
			}
		}
		return SyllableValidator.Validate(rawWord.ToLowerInvariant()) != SyllableValidationResult.ValidSyllable;
	}

	private static string GetAdjacentWord(List<TextToken> tokens, int currentIndex, int direction)
	{
		for (int i = currentIndex + direction; i >= 0 && i < tokens.Count; i += direction)
		{
			if (tokens[i].Type == TokenType.Word)
			{
				return tokens[i].Text.ToLowerInvariant();
			}
		}
		return null;
	}

	private static string ExtractContextSnippet(string text, int start, int end)
	{
		if (string.IsNullOrEmpty(text))
		{
			return string.Empty;
		}
		int num = Math.Max(0, start - 25);
		int num2 = Math.Min(text.Length, end + 25);
		string text2 = ((num > 0) ? "..." : string.Empty);
		string text3 = ((num2 < text.Length) ? "..." : string.Empty);
		string text4 = text.Substring(num, num2 - num).Replace('\r', ' ').Replace('\n', ' ');
		return text2 + text4.Trim() + text3;
	}

	private static bool IsAllUpper(string s)
	{
		if (string.IsNullOrEmpty(s))
		{
			return false;
		}
		foreach (char c in s)
		{
			if (char.IsLetter(c) && !char.IsUpper(c) && c != 'Đ')
			{
				return false;
			}
		}
		return true;
	}

	private static bool HasVietnameseVowel(string s)
	{
		if (string.IsNullOrEmpty(s))
		{
			return false;
		}
		for (int i = 0; i < s.Length; i++)
		{
			if (ViVowels.Contains(s[i]))
			{
				return true;
			}
		}
		return false;
	}

	private static bool IsRomanNumeral(string s)
	{
		if (string.IsNullOrEmpty(s) || s.Length > 15)
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
		return RomanNumeralRegex.IsMatch(text);
	}

	private static bool IsCommonAsciiAcronym(string s)
	{
		if (string.IsNullOrEmpty(s) || s.Length < 2 || s.Length > 6)
		{
			return false;
		}
		foreach (char c in s)
		{
			if (c < 'A' || c > 'Z')
			{
				return false;
			}
		}
		return true;
	}
}
