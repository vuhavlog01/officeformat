using System;
using System.Collections.Generic;
using System.Threading;
using VFormat.Core.SpellCheck.Processing;
using VFormat.Core.SpellCheck.Validation;

namespace VFormat.Core.SpellCheck.Engine;

public class CandidateGenerator
{
	private readonly SpellDictionaryManager _dictManager;

	public CandidateGenerator(SpellDictionaryManager dictManager)
	{
		_dictManager = dictManager;
	}

	public List<string> GenerateCandidates(string token, string prevWord = null, string nextWord = null)
	{
		return GenerateCandidates(token, prevWord, nextWord, CancellationToken.None);
	}

	public List<string> GenerateCandidates(string token, string prevWord, string nextWord, CancellationToken cancellationToken)
	{
		cancellationToken.ThrowIfCancellationRequested();
		HashSet<string> hashSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
		if (string.IsNullOrEmpty(token))
		{
			return new List<string>();
		}
		string text = VietnameseUnicodeNormalizer.NormalizeNfc(token).Trim();
		string text2 = text.ToLowerInvariant();
		if (_dictManager.TryGetTypoRule(text, out var entry))
		{
			for (int i = 0; i < entry.suggestions.Count; i++)
			{
				hashSet.Add(entry.suggestions[i]);
			}
			return new List<string>(hashSet);
		}
		if (!TonePlacementValidator.IsTonePlacedProperly(text, out var correctedWord) && _dictManager.ContainsWord(correctedWord))
		{
			hashSet.Add(correctedWord);
		}
		TonePlacementValidator.ExtractTone(text2, out var baseWord, out var _);
		string text3 = baseWord.Replace('đ', 'd');
		IEnumerable<string> candidatePool = _dictManager.GetCandidatePool(text2.Length);
		int num = 0;
		foreach (string item in candidatePool)
		{
			if ((num++ & 0x7F) == 0)
			{
				cancellationToken.ThrowIfCancellationRequested();
			}
			string text4 = item.ToLowerInvariant();
			TonePlacementValidator.ExtractTone(text4, out var baseWord2, out var _);
			string value = baseWord2.Replace('đ', 'd');
			if (text3.Equals(value, StringComparison.OrdinalIgnoreCase))
			{
				hashSet.Add(item);
				continue;
			}
			int num2 = ComputeDamerauLevenshtein(text2, text4);
			if (num2 <= 1)
			{
				hashSet.Add(item);
			}
			else if (num2 == 2 && text2.Length >= 4)
			{
				hashSet.Add(item);
			}
		}
		return new List<string>(hashSet);
	}

	public static int ComputeDamerauLevenshtein(string source, string target)
	{
		if (string.IsNullOrEmpty(source))
		{
			return target?.Length ?? 0;
		}
		if (string.IsNullOrEmpty(target))
		{
			return source.Length;
		}
		int length = source.Length;
		int length2 = target.Length;
		int[,] array = new int[length + 1, length2 + 1];
		for (int i = 0; i <= length; i++)
		{
			array[i, 0] = i;
		}
		for (int j = 0; j <= length2; j++)
		{
			array[0, j] = j;
		}
		for (int k = 1; k <= length; k++)
		{
			for (int l = 1; l <= length2; l++)
			{
				int num = ((source[k - 1] != target[l - 1]) ? 1 : 0);
				int val = array[k - 1, l] + 1;
				int val2 = array[k, l - 1] + 1;
				int val3 = array[k - 1, l - 1] + num;
				int num2 = Math.Min(val, Math.Min(val2, val3));
				if (k > 1 && l > 1 && source[k - 1] == target[l - 2] && source[k - 2] == target[l - 1])
				{
					num2 = Math.Min(num2, array[k - 2, l - 2] + num);
				}
				array[k, l] = num2;
			}
		}
		return array[length, length2];
	}
}
