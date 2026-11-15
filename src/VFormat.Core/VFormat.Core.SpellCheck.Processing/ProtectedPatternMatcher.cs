using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace VFormat.Core.SpellCheck.Processing;

public class ProtectedPatternMatcher
{
	private static readonly TimeSpan RegexTimeout = TimeSpan.FromMilliseconds(250.0);

	private static readonly Regex QuotedText = BoundedRegex("“[^”\\r\\n]{1,100}”|‘[^’\\r\\n]{1,100}’|'[^'\\r\\n]{1,100}'", RegexOptions.Compiled);

	private static readonly Regex MetaMentionBefore = BoundedRegex("(?:^|[\\s,:;])(?:cụm\\s+từ|từ|ví\\s+dụ|cách\\s+viết|thuật\\s+ngữ|tên\\s+gọi|tên\\s+sản\\s+phẩm|thương\\s+hiệu|nhan\\s+đề|tiêu\\s+đề)\\s*[:：]?\\s*$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private static readonly Regex MetaMentionAfter = BoundedRegex("^\\s*(?:là|được|chỉ\\s+là)\\s+(?:một\\s+)?(?:lỗi|cách\\s+viết\\s+sai|ví\\s+dụ|nêu|trích|giữ\\s+nguyên|nhan\\s+đề|tên\\s+sản\\s+phẩm|tên\\s+thương\\s+hiệu)", RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private readonly List<Regex> _patternRegexes;

	private readonly List<string> _protectedPhrases;

	private static Regex BoundedRegex(string pattern, RegexOptions options)
	{
		return new Regex(pattern, options, RegexTimeout);
	}

	public ProtectedPatternMatcher()
	{
		_patternRegexes = new List<Regex>();
		_protectedPhrases = new List<string>();
		InitializeDefaultPatterns();
	}

	private void InitializeDefaultPatterns()
	{
		_patternRegexes.Add(BoundedRegex("\\b\\d{1,5}/\\d{4}/[A-ZĐa-zđ0-9.-]+\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("\\b\\d{1,5}/[A-ZĐa-zđ0-9.-]+\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("\\b\\d{1,5}-[A-ZĐa-zđ0-9.-]+/[A-ZĐa-zđ0-9.-]+\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("(?:Số\\s*:\\s*)?[0-9.…—_]*/[A-ZĐa-zđ0-9.-]+\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("\\b[A-ZĐ]{2,}(?:-[A-ZĐ]{2,})+\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("\\b\\d{1,2}[-/]\\d{1,2}[-/]\\d{4}\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("(?i)ngày\\s+\\d{1,2}\\s+tháng\\s+\\d{1,2}\\s+năm\\s+\\d{4}\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("\\b\\d{1,3}(?:\\s\\d{3})+\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("\\b\\d+,\\d+\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("\\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("\\bhttps?://[^\\s/$.?#].[^\\s]*\\b", RegexOptions.Compiled));
		_patternRegexes.Add(BoundedRegex("\\b[A-ZĐ]{1,4}\\d{1,5}[A-ZĐ0-9]*\\b", RegexOptions.Compiled));
		_protectedPhrases.Add("Cộng hòa xã hội chủ nghĩa Việt Nam");
		_protectedPhrases.Add("Độc lập - Tự do - Hạnh phúc");
		_protectedPhrases.Add("ĐẢNG CỘNG SẢN VIỆT NAM");
		_protectedPhrases.Add("Quân đội nhân dân Việt Nam");
		_protectedPhrases.Add("Quân ủy Trung ương");
		_protectedPhrases.Add("Bộ Tổng Tham mưu");
		_protectedPhrases.Add("Tổng cục Chính trị");
		_protectedPhrases.Add("Bộ CHQS");
		_protectedPhrases.Add("Bộ Chỉ huy Quân sự");
		_protectedPhrases.Add("Ban CHQS");
		_protectedPhrases.Add("Ban Chỉ huy Quân sự");
		_protectedPhrases.Add("Sẵn sàng chiến đấu");
		_protectedPhrases.Add("Bộ Tư lệnh");
		_protectedPhrases.Add("Ủy ban nhân dân");
		_protectedPhrases.Add("Hội đồng nhân dân");
		_protectedPhrases.Add("Thủ tướng Chính phủ");
		_protectedPhrases.Add("Thủ đô Hà Nội");
		_protectedPhrases.Add("Thành phố Hồ Chí Minh");
		_protectedPhrases.Add("Bộ Giáo dục và Đào tạo");
		_protectedPhrases.Add("Bộ Quốc phòng");
		_protectedPhrases.Add("Bộ Tư pháp");
		_protectedPhrases.Add("Bộ Nội vụ");
		_protectedPhrases.Add("Văn phòng Trung ương Đảng");
		_protectedPhrases.Add("Ủy ban Kiểm tra Trung ương");
	}

	public void AddPhrase(string phrase)
	{
		if (!string.IsNullOrEmpty(phrase) && !_protectedPhrases.Contains(phrase))
		{
			_protectedPhrases.Add(phrase);
		}
	}

	public List<ProtectedSpan> FindProtectedSpans(string text)
	{
		List<ProtectedSpan> list = new List<ProtectedSpan>();
		if (string.IsNullOrEmpty(text))
		{
			return list;
		}
		foreach (Match item in QuotedText.Matches(text))
		{
			string text2 = item.Value.Substring(1, item.Value.Length - 2);
			if (text2.Length <= 35 && text2.IndexOfAny(new char[5] { '.', '?', '!', ';', ':' }) < 0 && text2.Split(new char[2] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).Length <= 4)
			{
				int num = Math.Max(0, item.Index - 48);
				string input = text.Substring(num, item.Index - num);
				int num2 = item.Index + item.Length;
				string input2 = text.Substring(num2, Math.Min(64, text.Length - num2));
				if (MetaMentionBefore.IsMatch(input) || MetaMentionAfter.IsMatch(input2))
				{
					AddNonOverlappingSpan(list, new ProtectedSpan(item.Index, num2, "QuotedExample", item.Value));
				}
			}
		}
		for (int i = 0; i < _protectedPhrases.Count; i++)
		{
			string text3 = _protectedPhrases[i];
			int startIndex = 0;
			while ((startIndex = text.IndexOf(text3, startIndex, StringComparison.OrdinalIgnoreCase)) >= 0)
			{
				int num3 = startIndex + text3.Length;
				AddNonOverlappingSpan(list, new ProtectedSpan(startIndex, num3, "ProtectedPhrase", text.Substring(startIndex, text3.Length)));
				startIndex = num3;
			}
		}
		for (int j = 0; j < _patternRegexes.Count; j++)
		{
			MatchCollection matchCollection = _patternRegexes[j].Matches(text);
			foreach (Match item2 in matchCollection)
			{
				AddNonOverlappingSpan(list, new ProtectedSpan(item2.Index, item2.Index + item2.Length, "ProtectedPattern", item2.Value));
			}
		}
		list.Sort((ProtectedSpan a, ProtectedSpan b) => a.Start.CompareTo(b.Start));
		return list;
	}

	private static void AddNonOverlappingSpan(List<ProtectedSpan> spans, ProtectedSpan candidate)
	{
		for (int i = 0; i < spans.Count; i++)
		{
			ProtectedSpan protectedSpan = spans[i];
			if (candidate.Start < protectedSpan.End && candidate.End > protectedSpan.Start)
			{
				if (candidate.End - candidate.Start > protectedSpan.End - protectedSpan.Start)
				{
					spans[i] = candidate;
				}
				return;
			}
		}
		spans.Add(candidate);
	}
}
