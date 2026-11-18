using System;
using System.Collections.Generic;
using VFormat.Core.SpellCheck.Models;
using VFormat.Core.SpellCheck.Processing;
using VFormat.Core.SpellCheck.Validation;

namespace VFormat.Core.SpellCheck.Engine;

public class CandidateRanker
{
	private readonly SpellDictionaryManager _dictManager;

	public CandidateRanker(SpellDictionaryManager dictManager)
	{
		_dictManager = dictManager;
	}

	public List<SpellSuggestion> RankCandidates(string originalToken, List<string> candidates, string prevWord = null, string nextWord = null)
	{
		List<SpellSuggestion> list = new List<SpellSuggestion>();
		List<KeyValuePair<double, SpellSuggestion>> list2 = new List<KeyValuePair<double, SpellSuggestion>>();
		if (candidates == null || candidates.Count == 0)
		{
			return list;
		}
		string text = VietnameseUnicodeNormalizer.NormalizeNfc(originalToken).Trim();
		string text2 = text.ToLowerInvariant();
		int num = TonePlacementValidator.ExtractTone(text2, out var baseWord, out var tonePos);
		string text3 = baseWord.Replace('đ', 'd');
		bool flag = _dictManager.TryGetTypoRule(text, out var entry);
		for (int i = 0; i < candidates.Count; i++)
		{
			string text4 = candidates[i];
			string text5 = text4.ToLowerInvariant();
			double num2 = 0.0;
			string reason = "Khoảng cách ký tự gần nhất";
			if (flag && entry.suggestions != null && entry.suggestions.Contains(text5))
			{
				num2 += 100.0;
				reason = "Quy tắc lỗi gõ phím chuẩn xác";
			}
			int num3 = TonePlacementValidator.ExtractTone(text5, out var baseWord2, out var tonePos2);
			string value = baseWord2.Replace('đ', 'd');
			if (text3.Equals(value, StringComparison.OrdinalIgnoreCase))
			{
				num2 += 35.0;
				reason = "Trùng gốc phụ âm/nguyên âm tiếng Việt";
			}
			if (num > 0 && num == num3 && tonePos != tonePos2 && string.Equals(baseWord, baseWord2, StringComparison.Ordinal))
			{
				num2 += 40.0;
				reason = "Sai vị trí đặt dấu thanh";
			}
			if (text2.Length > 0 && text5.Length > 0)
			{
				if (text2[0] == text5[0])
				{
					num2 += 5.0;
				}
				if (text2[text2.Length - 1] == text5[text5.Length - 1])
				{
					num2 += 5.0;
				}
			}
			if (_dictManager.IsAdminLegal(text5))
			{
				num2 += 15.0;
			}
			int wordFrequency = _dictManager.GetWordFrequency(text5);
			if (wordFrequency > 0)
			{
				num2 += Math.Min(20.0, Math.Log10(wordFrequency + 1) * 5.0);
			}
			if (!string.IsNullOrEmpty(prevWord))
			{
				int bigramFrequency = _dictManager.GetBigramFrequency(prevWord, text5);
				if (bigramFrequency > 0)
				{
					num2 += Math.Min(30.0, 10.0 + Math.Log10(bigramFrequency + 1) * 10.0);
					reason = "Phù hợp với ngữ cảnh từ liền trước";
				}
			}
			if (!string.IsNullOrEmpty(nextWord))
			{
				int bigramFrequency2 = _dictManager.GetBigramFrequency(text5, nextWord);
				if (bigramFrequency2 > 0)
				{
					num2 += Math.Min(30.0, 10.0 + Math.Log10(bigramFrequency2 + 1) * 10.0);
					reason = "Phù hợp với ngữ cảnh từ liền sau";
				}
			}
			int num4 = CandidateGenerator.ComputeDamerauLevenshtein(text2, text5);
			num2 -= (double)num4 * 10.0;
			string text6 = MatchCase(originalToken, text4);
			double score = 0.6;
			if (num2 >= 90.0)
			{
				score = 1.0;
			}
			else if (num2 >= 60.0)
			{
				score = 0.95;
			}
			else if (num2 >= 40.0)
			{
				score = 0.85;
			}
			else if (num2 >= 20.0)
			{
				score = 0.75;
			}
			list2.Add(new KeyValuePair<double, SpellSuggestion>(num2, new SpellSuggestion(text6, score, reason)));
		}
		list2.Sort((KeyValuePair<double, SpellSuggestion> a, KeyValuePair<double, SpellSuggestion> b) =>
		{
			int num6 = b.Key.CompareTo(a.Key);
			return (num6 == 0) ? StringComparer.OrdinalIgnoreCase.Compare(a.Value.Text, b.Value.Text) : num6;
		});
		for (int num5 = 0; num5 < list2.Count && num5 < 5; num5++)
		{
			list.Add(list2[num5].Value);
		}
		return list;
	}

	private static string MatchCase(string original, string target)
	{
		if (string.IsNullOrEmpty(original) || string.IsNullOrEmpty(target))
		{
			return target;
		}
		bool flag = true;
		for (int i = 0; i < original.Length; i++)
		{
			if (char.IsLetter(original[i]) && !char.IsUpper(original[i]))
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
			if (target.Length == 1)
			{
				return target.ToUpperInvariant();
			}
			return char.ToUpperInvariant(target[0]) + target.Substring(1);
		}
		return target.ToLowerInvariant();
	}
}
