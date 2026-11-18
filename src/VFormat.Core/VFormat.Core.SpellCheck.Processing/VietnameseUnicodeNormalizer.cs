using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace VFormat.Core.SpellCheck.Processing;

public static class VietnameseUnicodeNormalizer
{
	private static readonly Dictionary<string, string> ToneVariantMap = new Dictionary<string, string>(StringComparer.Ordinal)
	{
		{ "úy", "uý" },
		{ "uý", "úy" },
		{ "ùy", "uỳ" },
		{ "uỳ", "ùy" },
		{ "ủy", "uỷ" },
		{ "uỷ", "ủy" },
		{ "ũy", "uỹ" },
		{ "uỹ", "ũy" },
		{ "ụy", "uỵ" },
		{ "uỵ", "ụy" },
		{ "óa", "oá" },
		{ "oá", "óa" },
		{ "òa", "oà" },
		{ "oà", "òa" },
		{ "ỏa", "oả" },
		{ "oả", "ỏa" },
		{ "õa", "oã" },
		{ "oã", "õa" },
		{ "ọa", "oạ" },
		{ "oạ", "ọa" },
		{ "óe", "oé" },
		{ "oé", "óe" },
		{ "òe", "oè" },
		{ "oè", "òe" },
		{ "ỏe", "oẻ" },
		{ "oẻ", "ỏe" },
		{ "õe", "oẽ" },
		{ "oẽ", "õe" },
		{ "ọe", "oẹ" },
		{ "oẹ", "ọe" }
	};

	public static NormalizedMapping NormalizeWithMapping(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return new NormalizedMapping(string.Empty, string.Empty, new int[0], new int[0]);
		}
		if (text.IsNormalized(NormalizationForm.FormC))
		{
			int length = text.Length;
			int[] array = new int[length + 1];
			int[] array2 = new int[length + 1];
			for (int i = 0; i <= length; i++)
			{
				array[i] = i;
				array2[i] = i;
			}
			return new NormalizedMapping(text, text, array, array2);
		}
		TextElementEnumerator textElementEnumerator = StringInfo.GetTextElementEnumerator(text);
		StringBuilder stringBuilder = new StringBuilder();
		List<int> list = new List<int>();
		List<int> list2 = new List<int>();
		while (textElementEnumerator.MoveNext())
		{
			string textElement = textElementEnumerator.GetTextElement();
			int elementIndex = textElementEnumerator.ElementIndex;
			int item = elementIndex + textElement.Length;
			string text2 = textElement.Normalize(NormalizationForm.FormC);
			for (int j = 0; j < text2.Length; j++)
			{
				list.Add(elementIndex);
				list2.Add(item);
			}
			stringBuilder.Append(text2);
		}
		string normalizedText = stringBuilder.ToString();
		list.Add(text.Length);
		list2.Add(text.Length);
		return new NormalizedMapping(text, normalizedText, list.ToArray(), list2.ToArray());
	}

	public static string NormalizeNfc(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return string.Empty;
		}
		return text.Normalize(NormalizationForm.FormC);
	}

	public static string NormalizeQuotes(string text)
	{
		if (string.IsNullOrEmpty(text))
		{
			return string.Empty;
		}
		return text.Replace('’', '\'').Replace('‘', '\'');
	}

	public static string GetAlternativeTonePlacement(string word)
	{
		if (string.IsNullOrEmpty(word))
		{
			return null;
		}
		string text = word.ToLowerInvariant();
		foreach (KeyValuePair<string, string> item in ToneVariantMap)
		{
			int num = text.IndexOf(item.Key, StringComparison.Ordinal);
			if ((num > 0 && text[num - 1] == 'q') || num < 0)
			{
				continue;
			}
			string text2 = text.Substring(0, num) + item.Value + text.Substring(num + item.Key.Length);
			if (char.IsUpper(word[0]))
			{
				if (word.Length > 1 && char.IsUpper(word[1]))
				{
					return text2.ToUpperInvariant();
				}
				return char.ToUpperInvariant(text2[0]) + ((text2.Length > 1) ? text2.Substring(1) : string.Empty);
			}
			return text2;
		}
		return null;
	}
}
