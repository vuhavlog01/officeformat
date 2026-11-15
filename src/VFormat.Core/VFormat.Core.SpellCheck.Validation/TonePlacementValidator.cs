using System.Collections.Generic;
using System.Text;

namespace VFormat.Core.SpellCheck.Validation;

public static class TonePlacementValidator
{
	private static readonly Dictionary<char, KeyValuePair<char, int>> ToneMap;

	private static readonly Dictionary<string, char> ReverseToneMap;

	static TonePlacementValidator()
	{
		ToneMap = new Dictionary<char, KeyValuePair<char, int>>();
		ReverseToneMap = new Dictionary<string, char>();
		RegisterVowelGroup('a', "aàáảãạ");
		RegisterVowelGroup('ă', "ăằắẳẵặ");
		RegisterVowelGroup('â', "âầấẩẫậ");
		RegisterVowelGroup('e', "eèéẻẽẹ");
		RegisterVowelGroup('ê', "êềếểễệ");
		RegisterVowelGroup('i', "iìíỉĩị");
		RegisterVowelGroup('o', "oòóỏõọ");
		RegisterVowelGroup('ô', "ôồốổỗộ");
		RegisterVowelGroup('ơ', "ơờớởỡợ");
		RegisterVowelGroup('u', "uùúủũụ");
		RegisterVowelGroup('ư', "ưừứửữự");
		RegisterVowelGroup('y', "yỳýỷỹỵ");
	}

	private static void RegisterVowelGroup(char baseChar, string group)
	{
		for (int i = 0; i < group.Length; i++)
		{
			char c = group[i];
			ToneMap[c] = new KeyValuePair<char, int>(baseChar, i);
			ReverseToneMap[$"{baseChar}_{i}"] = c;
		}
	}

	public static int ExtractTone(string word, out string baseWord, out int tonePos)
	{
		tonePos = -1;
		int result = 0;
		StringBuilder stringBuilder = new StringBuilder();
		for (int i = 0; i < word.Length; i++)
		{
			char c = word[i];
			char key = char.ToLowerInvariant(c);
			if (ToneMap.TryGetValue(key, out var value))
			{
				if (value.Value > 0)
				{
					result = value.Value;
					tonePos = i;
				}
				stringBuilder.Append(char.IsUpper(c) ? char.ToUpperInvariant(value.Key) : value.Key);
			}
			else
			{
				stringBuilder.Append(c);
			}
		}
		baseWord = stringBuilder.ToString();
		return result;
	}

	public static string ApplyTone(string baseWord, int tone, int vowelIndex)
	{
		if (tone <= 0 || vowelIndex < 0 || vowelIndex >= baseWord.Length)
		{
			return baseWord;
		}
		char c = baseWord[vowelIndex];
		char c2 = char.ToLowerInvariant(c);
		if (ReverseToneMap.TryGetValue($"{c2}_{tone}", out var value))
		{
			StringBuilder stringBuilder = new StringBuilder(baseWord);
			stringBuilder[vowelIndex] = (char.IsUpper(c) ? char.ToUpperInvariant(value) : value);
			return stringBuilder.ToString();
		}
		return baseWord;
	}

	public static bool IsTonePlacedProperly(string word, out string correctedWord)
	{
		correctedWord = word;
		if (string.IsNullOrEmpty(word))
		{
			return true;
		}
		int num = ExtractTone(word, out var baseWord, out var tonePos);
		if (num == 0 || tonePos == -1)
		{
			return true;
		}
		List<int> list = new List<int>();
		for (int i = 0; i < baseWord.Length; i++)
		{
			char key = char.ToLowerInvariant(baseWord[i]);
			if (ToneMap.ContainsKey(key))
			{
				list.Add(i);
			}
		}
		if (list.Count <= 1)
		{
			return true;
		}
		int num2 = list[list.Count - 1];
		bool flag = num2 < baseWord.Length - 1;
		int num3 = tonePos;
		if (list.Count == 2)
		{
			char c = char.ToLowerInvariant(baseWord[list[0]]);
			char c2 = char.ToLowerInvariant(baseWord[list[1]]);
			if (flag)
			{
				num3 = list[1];
			}
			else
			{
				if ((c != 'i' || c2 != 'a') && (c != 'u' || c2 != 'a') && (c != 'ư' || c2 != 'a'))
				{
					return true;
				}
				num3 = list[0];
			}
		}
		else if (list.Count == 3)
		{
			num3 = list[1];
		}
		if (tonePos != num3)
		{
			correctedWord = ApplyTone(baseWord, num, num3);
			return false;
		}
		return true;
	}
}
