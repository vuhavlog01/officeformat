using System;
using System.Collections.Generic;

namespace VFormat.Core.SpellCheck.Validation;

public static class SyllableValidator
{
	private static readonly HashSet<char> Vowels = new HashSet<char>("aàáảãạăằắẳẵặâầấẩẫậeèéẻẽẹêềếểễệiìíỉĩịoòóỏõọôồốổỗộơờớởỡợuùúủũụưừứửữựyỳýỷỹỵ");

	private static readonly string[] ValidOnsets = new string[28]
	{
		"ngh", "ng", "nh", "ch", "gh", "gi", "kh", "ph", "qu", "th",
		"tr", "b", "c", "d", "đ", "g", "h", "k", "l", "m",
		"n", "p", "r", "s", "t", "v", "x", ""
	};

	private static readonly string[] ValidCodas = new string[9] { "ng", "nh", "ch", "c", "m", "n", "p", "t", "" };

	public static SyllableValidationResult Validate(string word)
	{
		if (string.IsNullOrEmpty(word))
		{
			return SyllableValidationResult.InvalidSyllable;
		}
		string text = word.ToLowerInvariant();
		foreach (char c in text)
		{
			if (c == 'f' || c == 'j' || c == 'w' || c == 'z')
			{
				return SyllableValidationResult.ForeignOrUnknown;
			}
		}
		bool flag = false;
		for (int j = 0; j < text.Length; j++)
		{
			if (Vowels.Contains(text[j]))
			{
				flag = true;
				break;
			}
		}
		if (!flag)
		{
			return SyllableValidationResult.InvalidSyllable;
		}
		string text2 = null;
		for (int k = 0; k < ValidOnsets.Length; k++)
		{
			string text3 = ValidOnsets[k];
			if (text3.Length <= 0 || !text.StartsWith(text3, StringComparison.Ordinal))
			{
				continue;
			}
			string text4 = text.Substring(text3.Length);
			bool flag2 = false;
			for (int l = 0; l < text4.Length; l++)
			{
				if (Vowels.Contains(text4[l]))
				{
					flag2 = true;
					break;
				}
			}
			if (flag2)
			{
				text2 = text3;
				break;
			}
		}
		if (text2 == null)
		{
			text2 = "";
		}
		string text5 = text.Substring(text2.Length);
		string text6 = "";
		for (int m = 0; m < ValidCodas.Length; m++)
		{
			string text7 = ValidCodas[m];
			if (text7.Length > 0 && text5.EndsWith(text7, StringComparison.Ordinal))
			{
				text6 = text7;
				break;
			}
		}
		string text8 = ((text6.Length > 0) ? text5.Substring(0, text5.Length - text6.Length) : text5);
		if (text8.Length == 0 || text8.Length > 3)
		{
			return SyllableValidationResult.InvalidSyllable;
		}
		for (int n = 0; n < text8.Length; n++)
		{
			if (!Vowels.Contains(text8[n]))
			{
				return SyllableValidationResult.InvalidSyllable;
			}
		}
		return SyllableValidationResult.ValidSyllable;
	}
}
