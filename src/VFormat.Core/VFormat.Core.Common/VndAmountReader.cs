using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace VFormat.Core.Common;

public static class VndAmountReader
{
	public const int MaxDigits = 18;

	private static readonly string[] Digits = new string[10] { "không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín" };

	private static readonly Regex CurrencySuffix = new Regex("\\s*(vnđ|vnd|đồng|đ|₫)\\s*\\.?\\s*$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

	private static readonly Regex DotGrouped = new Regex("^\\d{1,3}(?:\\.\\d{3})+(?:,(\\d+))?$");

	private static readonly Regex CommaGrouped = new Regex("^\\d{1,3}(?:,\\d{3})+(?:\\.(\\d+))?$");

	private static readonly Regex SpaceGrouped = new Regex("^\\d{1,3}(?: \\d{3})+$");

	private static readonly Regex Plain = new Regex("^\\d+(?:[.,](\\d+))?$");

	public static bool TryConvert(string text, out string words, out string error)
	{
		return TryConvert(text, null, out words, out error);
	}

	public static bool TryConvert(string text, VndReadingOptions options, out string words, out string error)
	{
		words = null;
		if (!TryParse(text, out var digits, out var negative, out error))
		{
			return false;
		}
		words = ToWords(digits, negative, options);
		return true;
	}

	public static bool TryParse(string text, out string digits, out bool negative, out string error)
	{
		digits = null;
		negative = false;
		error = null;
		string input = (text ?? string.Empty).Replace('\u00a0', ' ').Replace('\u202f', ' ').Trim();
		input = CurrencySuffix.Replace(input, string.Empty).Trim();
		if (input.StartsWith("-", StringComparison.Ordinal) || input.StartsWith("−", StringComparison.Ordinal))
		{
			negative = true;
			input = input.Substring(1).Trim();
		}
		Match match = DotGrouped.Match(input);
		if (!match.Success)
		{
			match = CommaGrouped.Match(input);
		}
		if (!match.Success)
		{
			match = SpaceGrouped.Match(input);
		}
		if (!match.Success)
		{
			match = Plain.Match(input);
		}
		if (!match.Success)
		{
			error = "Vùng chọn không phải là một số tiền (ví dụ hợp lệ: 2345, 2.345.000, 2.345.000 đồng).";
			return false;
		}
		string text2 = input;
		if (match.Groups.Count > 1 && match.Groups[1].Success)
		{
			string value = match.Groups[1].Value;
			if (value.Trim('0').Length > 0)
			{
				error = "Số tiền VND không có phần lẻ thập phân (ví dụ 2.345,50). Vui lòng làm tròn trước khi chuyển.";
				return false;
			}
			text2 = input.Substring(0, input.Length - value.Length - 1);
		}
		StringBuilder stringBuilder = new StringBuilder(text2.Length);
		string text3 = text2;
		foreach (char c in text3)
		{
			if (c >= '0' && c <= '9')
			{
				stringBuilder.Append(c);
			}
		}
		string text4 = stringBuilder.ToString().TrimStart('0');
		if (text4.Length == 0)
		{
			text4 = "0";
		}
		if (text4.Length > 18)
		{
			error = "Số tiền quá lớn (tối đa 18 chữ số).";
			return false;
		}
		digits = text4;
		return true;
	}

	public static string ToWords(string digits, bool negative)
	{
		return ToWords(digits, negative, null);
	}

	public static string ToWords(string digits, bool negative, VndReadingOptions options)
	{
		VndReadingOptions vndReadingOptions = options ?? VndReadingOptions.Default;
		string text = (digits ?? "0").TrimStart('0');
		string text2;
		if (text.Length == 0)
		{
			text2 = "không";
		}
		else
		{
			text2 = Read(text, full: false, vndReadingOptions);
			if (negative)
			{
				text2 = "âm " + text2;
			}
		}
		if (vndReadingOptions.CapitalizeFirst)
		{
			text2 = char.ToUpper(text2[0]) + text2.Substring(1);
		}
		return text2 + " đồng";
	}

	private static string Read(string d, bool full, VndReadingOptions o)
	{
		d = d.TrimStart('0');
		if (d.Length == 0)
		{
			d = "0";
		}
		if (d.Length > 9)
		{
			string d2 = d.Substring(0, d.Length - 9);
			string text = d.Substring(d.Length - 9);
			string text2 = Read(d2, full, o) + " tỷ";
			if (text.Trim('0').Length > 0)
			{
				text2 = text2 + " " + ReadUpToBillion(text, full: true, o);
			}
			return text2;
		}
		return ReadUpToBillion(d, full, o);
	}

	private static string ReadUpToBillion(string d, bool full, VndReadingOptions o)
	{
		string[] array = new string[3]
		{
			"",
			o.ThousandAsNgan ? " ngàn" : " nghìn",
			" triệu"
		};
		int totalWidth = (d.Length + 2) / 3 * 3;
		d = d.PadLeft(totalWidth, '0');
		int num = d.Length / 3;
		List<string> list = new List<string>();
		bool full2 = full;
		for (int i = 0; i < num; i++)
		{
			string text = d.Substring(i * 3, 3);
			if (!(text == "000"))
			{
				list.Add(ReadGroup(text[0] - 48, text[1] - 48, text[2] - 48, full2, o) + array[num - 1 - i]);
				full2 = true;
			}
		}
		return string.Join(" ", list.ToArray());
	}

	private static string ReadGroup(int h, int t, int u, bool full, VndReadingOptions o)
	{
		List<string> list = new List<string>();
		if (h > 0 || full)
		{
			list.Add(Digits[h] + " trăm");
		}
		switch (t)
		{
		case 0:
			if (u > 0)
			{
				if (h > 0 || full)
				{
					list.Add(o.ZeroTensAsLe ? "lẻ" : "linh");
				}
				list.Add(Digits[u]);
			}
			break;
		case 1:
			list.Add("mười");
			if (u > 0)
			{
				list.Add((u == 5) ? "lăm" : Digits[u]);
			}
			break;
		default:
			list.Add(Digits[t] + " mươi");
			if (u == 1)
			{
				list.Add("mốt");
			}
			else if (u == 4 && o.FourAsTu)
			{
				list.Add("tư");
			}
			else if (u == 5)
			{
				list.Add("lăm");
			}
			else if (u > 0)
			{
				list.Add(Digits[u]);
			}
			break;
		}
		return string.Join(" ", list.ToArray());
	}
}
