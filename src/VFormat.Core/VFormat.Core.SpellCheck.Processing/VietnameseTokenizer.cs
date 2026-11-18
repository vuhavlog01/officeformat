using System.Collections.Generic;
using VFormat.Core.SpellCheck.Models;

namespace VFormat.Core.SpellCheck.Processing;

public class VietnameseTokenizer
{
	private readonly ProtectedPatternMatcher _protectedMatcher;

	public VietnameseTokenizer(ProtectedPatternMatcher protectedMatcher = null)
	{
		_protectedMatcher = protectedMatcher ?? new ProtectedPatternMatcher();
	}

	public static bool IsVietnameseLetter(char c)
	{
		if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z'))
		{
			return true;
		}
		if (c == 'đ' || c == 'Đ')
		{
			return true;
		}
		char c2 = c;
		if (c2 <= 'ĩ')
		{
			switch (c2)
			{
			case 'À':
			case 'Á':
			case 'Â':
			case 'Ã':
			case 'È':
			case 'É':
			case 'Ê':
			case 'Ì':
			case 'Í':
			case 'Ò':
			case 'Ó':
			case 'Ô':
			case 'Õ':
			case 'Ù':
			case 'Ú':
			case 'Ý':
			case 'à':
			case 'á':
			case 'â':
			case 'ã':
			case 'è':
			case 'é':
			case 'ê':
			case 'ì':
			case 'í':
			case 'ò':
			case 'ó':
			case 'ô':
			case 'õ':
			case 'ù':
			case 'ú':
			case 'ý':
			case 'Ă':
			case 'ă':
			case 'Ĩ':
			case 'ĩ':
				break;
			default:
				goto IL_0327;
			}
		}
		else
		{
			switch (c2)
			{
			case 'Ũ':
			case 'ũ':
			case 'Ơ':
			case 'ơ':
			case 'Ư':
			case 'ư':
			case 'Ạ':
			case 'ạ':
			case 'Ả':
			case 'ả':
			case 'Ấ':
			case 'ấ':
			case 'Ầ':
			case 'ầ':
			case 'Ẩ':
			case 'ẩ':
			case 'Ẫ':
			case 'ẫ':
			case 'Ậ':
			case 'ậ':
			case 'Ắ':
			case 'ắ':
			case 'Ằ':
			case 'ằ':
			case 'Ẳ':
			case 'ẳ':
			case 'Ẵ':
			case 'ẵ':
			case 'Ặ':
			case 'ặ':
			case 'Ẹ':
			case 'ẹ':
			case 'Ẻ':
			case 'ẻ':
			case 'Ẽ':
			case 'ẽ':
			case 'Ế':
			case 'ế':
			case 'Ề':
			case 'ề':
			case 'Ể':
			case 'ể':
			case 'Ễ':
			case 'ễ':
			case 'Ệ':
			case 'ệ':
			case 'Ỉ':
			case 'ỉ':
			case 'Ị':
			case 'ị':
			case 'Ọ':
			case 'ọ':
			case 'Ỏ':
			case 'ỏ':
			case 'Ố':
			case 'ố':
			case 'Ồ':
			case 'ồ':
			case 'Ổ':
			case 'ổ':
			case 'Ỗ':
			case 'ỗ':
			case 'Ộ':
			case 'ộ':
			case 'Ớ':
			case 'ớ':
			case 'Ờ':
			case 'ờ':
			case 'Ở':
			case 'ở':
			case 'Ỡ':
			case 'ỡ':
			case 'Ợ':
			case 'ợ':
			case 'Ụ':
			case 'ụ':
			case 'Ủ':
			case 'ủ':
			case 'Ứ':
			case 'ứ':
			case 'Ừ':
			case 'ừ':
			case 'Ử':
			case 'ử':
			case 'Ữ':
			case 'ữ':
			case 'Ự':
			case 'ự':
			case 'Ỳ':
			case 'ỳ':
			case 'Ỵ':
			case 'ỵ':
			case 'Ỷ':
			case 'ỷ':
			case 'Ỹ':
			case 'ỹ':
				break;
			default:
				goto IL_0327;
			}
		}
		return true;
		IL_0327:
		return false;
	}

	public List<TextToken> Tokenize(string text)
	{
		List<TextToken> list = new List<TextToken>();
		if (string.IsNullOrEmpty(text))
		{
			return list;
		}
		List<ProtectedSpan> list2 = _protectedMatcher.FindProtectedSpans(text);
		int num = 0;
		int length = text.Length;
		int i = 0;
		while (i < length)
		{
			if (num < list2.Count && i >= list2[num].Start && i < list2[num].End)
			{
				ProtectedSpan protectedSpan = list2[num];
				list.Add(new TextToken(text.Substring(protectedSpan.Start, protectedSpan.End - protectedSpan.Start), protectedSpan.Start, protectedSpan.End, TokenType.Protected));
				i = protectedSpan.End;
				num++;
				continue;
			}
			if (num < list2.Count && i >= list2[num].End)
			{
				num++;
				continue;
			}
			char c = text[i];
			if (char.IsWhiteSpace(c))
			{
				int num2 = i;
				for (; i < length && char.IsWhiteSpace(text[i]); i++)
				{
				}
				list.Add(new TextToken(text.Substring(num2, i - num2), num2, i, TokenType.Whitespace));
			}
			else if (char.IsDigit(c))
			{
				int num3 = i;
				for (; i < length && (char.IsDigit(text[i]) || text[i] == '.' || text[i] == ',') && ((text[i] != '.' && text[i] != ',') || (i + 1 < length && char.IsDigit(text[i + 1]))); i++)
				{
				}
				list.Add(new TextToken(text.Substring(num3, i - num3), num3, i, TokenType.Number));
			}
			else if (IsVietnameseLetter(c))
			{
				int num4 = i;
				while (i < length)
				{
					char c2 = text[i];
					if (IsVietnameseLetter(c2))
					{
						i++;
						continue;
					}
					if ((c2 != '\'' && c2 != '’' && c2 != '‘' && c2 != '-' && c2 != '/') || i + 1 >= length || !IsVietnameseLetter(text[i + 1]))
					{
						break;
					}
					i += 2;
				}
				list.Add(new TextToken(text.Substring(num4, i - num4), num4, i, TokenType.Word));
			}
			else
			{
				list.Add(new TextToken(c.ToString(), i, i + 1, TokenType.Punctuation));
				i++;
			}
		}
		return list;
	}
}
