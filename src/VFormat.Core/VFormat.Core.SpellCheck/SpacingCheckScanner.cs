using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Threading;
using VFormat.Core.SpellCheck.Models;

namespace VFormat.Core.SpellCheck;

public static class SpacingCheckScanner
{
	private struct SpanRange
	{
		public int Start;
		public int End;

		public SpanRange(int start, int end)
		{
			Start = start;
			End = end;
		}

		public bool Overlaps(int start, int end)
		{
			return start < End && end > Start;
		}
	}

	private static readonly Regex ProtectedBlankPattern = new Regex(
		@"(?:\bngày[ \u00A0]{2,20}tháng(?:[ \u00A0]{2,20}năm)?\b|\btháng[ \u00A0]{2,20}năm\b|\bSố[ \u00A0]*:?[ \u00A0]{2,20}[/\-]|\s\.{3,}\s|_{3,}|\.{4,})",
		RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private static readonly Regex LeadingSpacesPattern = new Regex(
		@"^([ \u00A0]+)(?=\S)",
		RegexOptions.Compiled);

	private static readonly Regex TrailingSpacesPattern = new Regex(
		@"(?<=\S)([ \u00A0]+)$",
		RegexOptions.Compiled);

	private static readonly Regex TabAfterMarkerPattern = new Regex(
		@"^(\s*(?:\d{1,2}(?:\.\d{1,2})*|[a-zđ]|[IVXLCDM]{1,5})[.)])([ \u00A0]*\t+[ \u00A0]*)(?=\S)",
		RegexOptions.IgnoreCase | RegexOptions.Compiled);

	private static readonly Regex SpaceAfterOpenBracketPattern = new Regex(
		@"([(\[{“‘])([ \u00A0]+)(?=\S)",
		RegexOptions.Compiled);

	private static readonly Regex SpaceBeforeCloseBracketPattern = new Regex(
		@"(?<=\S)([ \u00A0]+)([)\]}”’])",
		RegexOptions.Compiled);

	private static readonly Regex SpaceBeforePunctPattern = new Regex(
		@"(?<=\S)([ \u00A0]+)([,.;:?!]+)",
		RegexOptions.Compiled);

	private static readonly Regex MultiSpaceBetweenWordsPattern = new Regex(
		@"(?<=\S)([ \u00A0]{2,})(?=\S)",
		RegexOptions.Compiled);

	private static readonly Regex NbspBetweenWordsPattern = new Regex(
		@"(?<=\S)(\u00A0)(?=\S)",
		RegexOptions.Compiled);

	private static readonly Regex MissingSpaceAfterPunctPattern = new Regex(
		@"([\p{L}\d]+)([,;:])(?=[\p{L}])",
		RegexOptions.Compiled);

	private static readonly Regex StraightQuotePattern = new Regex(
		"\"",
		RegexOptions.Compiled);

	private static readonly Regex EnDashPattern = new Regex(
		@"[\u2013\u2014]",
		RegexOptions.Compiled);

	public static List<SpellCheckIssue> CheckParagraph(string text, int paragraphIndex, CancellationToken cancellationToken = default)
	{
		List<SpellCheckIssue> issues = new List<SpellCheckIssue>();
		if (string.IsNullOrEmpty(text))
		{
			return issues;
		}
		cancellationToken.ThrowIfCancellationRequested();

		List<SpanRange> protectedSpans = new List<SpanRange>();
		foreach (Match m in ProtectedBlankPattern.Matches(text))
		{
			if (m.Success)
			{
				protectedSpans.Add(new SpanRange(m.Index, m.Index + m.Length));
			}
		}

		List<SpanRange> claimedSpans = new List<SpanRange>();

		// 1. Khoảng trắng thừa ở đầu đoạn văn (Type = Spelling -> "Khoảng trắng")
		Match leadMatch = LeadingSpacesPattern.Match(text);
		if (leadMatch.Success)
		{
			string spaces = leadMatch.Groups[1].Value;
			int start = leadMatch.Index;
			int end = start + leadMatch.Length;
			claimedSpans.Add(new SpanRange(start, end));

			SpellCheckIssue issue = new SpellCheckIssue
			{
				Start = start,
				End = end,
				ParagraphIndex = paragraphIndex,
				Original = spaces,
				Type = SpellCheckIssueType.Spelling,
				Severity = SpellCheckSeverity.Warning,
				RuleId = "SPACE_LEADING",
				Source = "Quy tắc trình bày văn bản hành chính (NĐ 30/2020 & HD 05-HD/VPTW)",
				Description = $"Thừa {spaces.Length} khoảng trắng ở đầu đoạn văn. Không dùng phím Space để thụt đầu dòng.",
				ContextSnippet = BuildContextSnippet(text, start, end)
			};
			issue.Suggestions.Add(new SpellSuggestion(
				string.Empty,
				1.0,
				$"Xóa {spaces.Length} khoảng trắng thừa ở đầu đoạn"));
			issues.Add(issue);
		}

		// 2. Khoảng trắng thừa ở cuối đoạn văn (Type = Spelling -> "Khoảng trắng")
		Match trailMatch = TrailingSpacesPattern.Match(text);
		if (trailMatch.Success)
		{
			int start = trailMatch.Index;
			int end = start + trailMatch.Length;
			if (!OverlapsAny(claimedSpans, start, end))
			{
				string spaces = trailMatch.Groups[1].Value;
				claimedSpans.Add(new SpanRange(start, end));

				SpellCheckIssue issue = new SpellCheckIssue
				{
					Start = start,
					End = end,
					ParagraphIndex = paragraphIndex,
					Original = spaces,
					Type = SpellCheckIssueType.Spelling,
					Severity = SpellCheckSeverity.Info,
					RuleId = "SPACE_TRAILING",
					Source = "Quy tắc làm sạch văn bản OfficeFormat",
					Description = $"Thừa {spaces.Length} khoảng trắng vô hình ở cuối đoạn văn.",
					ContextSnippet = BuildContextSnippet(text, start, end)
				};
				issue.Suggestions.Add(new SpellSuggestion(
					string.Empty,
					1.0,
					$"Xóa {spaces.Length} khoảng trắng thừa ở cuối đoạn"));
				issues.Add(issue);
			}
		}

		// 3. Dấu Tab sau ký hiệu đề mục (Type = Spelling -> "Khoảng trắng")
		Match tabMatch = TabAfterMarkerPattern.Match(text);
		if (tabMatch.Success)
		{
			Group tabGroup = tabMatch.Groups[2];
			int start = tabGroup.Index;
			int end = start + tabGroup.Length;
			if (!OverlapsAny(claimedSpans, start, end))
			{
				claimedSpans.Add(new SpanRange(start, end));

				SpellCheckIssue issue = new SpellCheckIssue
				{
					Start = start,
					End = end,
					ParagraphIndex = paragraphIndex,
					Original = tabGroup.Value,
					Type = SpellCheckIssueType.Spelling,
					Severity = SpellCheckSeverity.Warning,
					RuleId = "SPACE_TAB_AFTER_MARKER",
					Source = "Quy tắc trình bày đề mục văn bản (NĐ 30/2020 & HD 05-HD/VPTW)",
					Description = "Sử dụng ký tự Tab sau dấu chấm đề mục gây lệch khoảng cách khi căn đều 2 lề (Justify).",
					ContextSnippet = BuildContextSnippet(text, start, end)
				};
				issue.Suggestions.Add(new SpellSuggestion(
					" ",
					1.0,
					"Thay ký tự Tab bằng 1 dấu cách chuẩn sau đề mục"));
				issues.Add(issue);
			}
		}

		// 4. Dấu cách thừa sau dấu mở ngoặc / mở nháy (Type = Capitalization -> "Dấu câu / Ngoặc")
		foreach (Match m in SpaceAfterOpenBracketPattern.Matches(text))
		{
			int start = m.Index;
			int end = start + m.Length;
			if (OverlapsAny(protectedSpans, start, end) || OverlapsAny(claimedSpans, start, end))
			{
				continue;
			}
			string openBracket = m.Groups[1].Value;
			string spaces = m.Groups[2].Value;
			claimedSpans.Add(new SpanRange(start, end));

			SpellCheckIssue issue = new SpellCheckIssue
			{
				Start = start,
				End = end,
				ParagraphIndex = paragraphIndex,
				Original = m.Value,
				Type = SpellCheckIssueType.Capitalization,
				Severity = SpellCheckSeverity.Error,
				RuleId = "SPACE_AFTER_OPEN_BRACKET",
				Source = "Quy tắc dấu câu tiếng Việt (NĐ 30/2020)",
				Description = $"Thừa {spaces.Length} dấu cách sau dấu mở '{openBracket}'. Chữ bên trong phải viết sát dấu mở ngoặc/nháy.",
				ContextSnippet = BuildContextSnippet(text, start, end)
			};
			issue.Suggestions.Add(new SpellSuggestion(
				openBracket,
				1.0,
				$"Xóa dấu cách thừa sau dấu '{openBracket}'"));
			issues.Add(issue);
		}

		// 5. Dấu cách thừa trước dấu đóng ngoặc / đóng nháy (Type = Capitalization -> "Dấu câu / Ngoặc")
		foreach (Match m in SpaceBeforeCloseBracketPattern.Matches(text))
		{
			int start = m.Index;
			int end = start + m.Length;
			if (OverlapsAny(protectedSpans, start, end) || OverlapsAny(claimedSpans, start, end))
			{
				continue;
			}
			string spaces = m.Groups[1].Value;
			string closeBracket = m.Groups[2].Value;
			claimedSpans.Add(new SpanRange(start, end));

			SpellCheckIssue issue = new SpellCheckIssue
			{
				Start = start,
				End = end,
				ParagraphIndex = paragraphIndex,
				Original = m.Value,
				Type = SpellCheckIssueType.Capitalization,
				Severity = SpellCheckSeverity.Error,
				RuleId = "SPACE_BEFORE_CLOSE_BRACKET",
				Source = "Quy tắc dấu câu tiếng Việt (NĐ 30/2020)",
				Description = $"Thừa {spaces.Length} dấu cách trước dấu đóng '{closeBracket}'. Dấu đóng ngoặc/nháy phải viết sát vào từ đứng trước.",
				ContextSnippet = BuildContextSnippet(text, start, end)
			};
			issue.Suggestions.Add(new SpellSuggestion(
				closeBracket,
				1.0,
				$"Xóa dấu cách thừa trước dấu '{closeBracket}'"));
			issues.Add(issue);
		}

		// 6. Dấu cách thừa trước dấu câu , . ; : ? ! (Type = Capitalization -> "Dấu câu / Ngoặc")
		foreach (Match m in SpaceBeforePunctPattern.Matches(text))
		{
			int start = m.Index;
			int end = start + m.Length;
			if (OverlapsAny(protectedSpans, start, end) || OverlapsAny(claimedSpans, start, end))
			{
				continue;
			}
			string spaces = m.Groups[1].Value;
			string punct = m.Groups[2].Value;
			if (punct.StartsWith("..."))
			{
				continue;
			}
			claimedSpans.Add(new SpanRange(start, end));

			SpellCheckIssue issue = new SpellCheckIssue
			{
				Start = start,
				End = end,
				ParagraphIndex = paragraphIndex,
				Original = m.Value,
				Type = SpellCheckIssueType.Capitalization,
				Severity = SpellCheckSeverity.Error,
				RuleId = "SPACE_BEFORE_PUNCT",
				Source = "Quy tắc dấu câu tiếng Việt (NĐ 30/2020)",
				Description = $"Thừa {spaces.Length} dấu cách trước dấu câu '{punct}'. Dấu câu phải đặt sát ngay sau từ đứng trước.",
				ContextSnippet = BuildContextSnippet(text, start, end)
			};
			issue.Suggestions.Add(new SpellSuggestion(
				punct,
				1.0,
				$"Xóa dấu cách thừa trước dấu '{punct}'"));
			issues.Add(issue);
		}

		// 7. Thiếu dấu cách sau dấu câu , ; : (Type = Capitalization -> "Dấu câu / Ngoặc")
		foreach (Match m in MissingSpaceAfterPunctPattern.Matches(text))
		{
			string leftWord = m.Groups[1].Value;
			Group punctGroup = m.Groups[2];
			int start = punctGroup.Index;
			int end = start + punctGroup.Length;
			if (OverlapsAny(protectedSpans, start, end) || OverlapsAny(claimedSpans, start, end))
			{
				continue;
			}
			if (punctGroup.Value == ":" &&
				(leftWord.Equals("http", StringComparison.OrdinalIgnoreCase) ||
				 leftWord.Equals("https", StringComparison.OrdinalIgnoreCase) ||
				 leftWord.Equals("mailto", StringComparison.OrdinalIgnoreCase)))
			{
				continue;
			}
			char punctChar = punctGroup.Value[0];
			claimedSpans.Add(new SpanRange(start, end));

			SpellCheckIssue issue = new SpellCheckIssue
			{
				Start = start,
				End = end,
				ParagraphIndex = paragraphIndex,
				Original = punctGroup.Value,
				Type = SpellCheckIssueType.Capitalization,
				Severity = SpellCheckSeverity.Warning,
				RuleId = "SPACE_MISSING_AFTER_PUNCT",
				Source = "Quy tắc dấu câu tiếng Việt (NĐ 30/2020)",
				Description = $"Thiếu dấu cách sau dấu câu '{punctChar}'. Cần có 1 dấu cách ngăn cách với từ tiếp theo.",
				ContextSnippet = BuildContextSnippet(text, start, end)
			};
			issue.Suggestions.Add(new SpellSuggestion(
				punctGroup.Value + " ",
				1.0,
				$"Thêm 1 dấu cách chuẩn sau dấu '{punctChar}'"));
			issues.Add(issue);
		}

		// 8. Nhiều dấu cách liên tiếp (2+ dấu cách) giữa các từ (Type = Spelling -> "Khoảng trắng")
		foreach (Match m in MultiSpaceBetweenWordsPattern.Matches(text))
		{
			int start = m.Index;
			int end = start + m.Length;
			if (OverlapsAny(protectedSpans, start, end) || OverlapsAny(claimedSpans, start, end))
			{
				continue;
			}
			string spaces = m.Groups[1].Value;
			claimedSpans.Add(new SpanRange(start, end));

			SpellCheckIssue issue = new SpellCheckIssue
			{
				Start = start,
				End = end,
				ParagraphIndex = paragraphIndex,
				Original = spaces,
				Type = SpellCheckIssueType.Spelling,
				Severity = SpellCheckSeverity.Error,
				RuleId = "SPACE_MULTIPLE",
				Source = "Quy tắc soạn thảo văn bản hành chính (NĐ 30/2020 & HD 05-HD/VPTW)",
				Description = $"Có {spaces.Length} khoảng trắng liên tiếp giữa các từ. Giữa các từ chỉ dùng duy nhất 1 dấu cách.",
				ContextSnippet = BuildContextSnippet(text, start, end)
			};
			issue.Suggestions.Add(new SpellSuggestion(
				" ",
				1.0,
				$"Rút gọn {spaces.Length} khoảng trắng thành 1 dấu cách chuẩn"));
			issue.Suggestions.Add(new SpellSuggestion(
				string.Empty,
				0.65,
				"Xóa hoàn toàn khoảng trắng (ghép sát 2 từ)"));
			issues.Add(issue);
		}

		// 9. Khoảng trắng cứng (Non-breaking space \u00A0) đơn lẻ (Type = Spelling -> "Khoảng trắng")
		foreach (Match m in NbspBetweenWordsPattern.Matches(text))
		{
			int start = m.Index;
			int end = start + m.Length;
			if (OverlapsAny(protectedSpans, start, end) || OverlapsAny(claimedSpans, start, end))
			{
				continue;
			}
			claimedSpans.Add(new SpanRange(start, end));

			SpellCheckIssue issue = new SpellCheckIssue
			{
				Start = start,
				End = end,
				ParagraphIndex = paragraphIndex,
				Original = m.Value,
				Type = SpellCheckIssueType.Spelling,
				Severity = SpellCheckSeverity.Info,
				RuleId = "SPACE_NBSP",
				Source = "Quy tắc chuẩn hóa ký tự Unicode OfficeFormat",
				Description = "Phát hiện ký tự khoảng trắng cứng (Non-breaking space), có thể làm giãn dòng bất thường khi căn đều 2 lề.",
				ContextSnippet = BuildContextSnippet(text, start, end)
			};
			issue.Suggestions.Add(new SpellSuggestion(
				" ",
				1.0,
				"Chuyển thành dấu cách thường (Standard Space)"));
			issues.Add(issue);
		}

		// 10. Dấu nháy thẳng " đổi sang “ (Type = Abbreviation -> "Dấu nháy (\")")
		foreach (Match m in StraightQuotePattern.Matches(text))
		{
			int start = m.Index;
			int end = start + m.Length;
			if (OverlapsAny(claimedSpans, start, end))
			{
				continue;
			}
			claimedSpans.Add(new SpanRange(start, end));

			SpellCheckIssue issue = new SpellCheckIssue
			{
				Start = start,
				End = end,
				ParagraphIndex = paragraphIndex,
				Original = "\"",
				Type = SpellCheckIssueType.Abbreviation,
				Severity = SpellCheckSeverity.Warning,
				RuleId = "SPACE_STRAIGHT_QUOTE",
				Source = "Quy tắc dấu câu & ký hiệu văn bản (NĐ 30/2020 & HD 05-HD/VPTW)",
				Description = "Phát hiện dấu nháy thẳng (\"). Theo quy chuẩn trình bày văn bản, nên đổi sang dấu nháy kép cong (“ hoặc ”).",
				ContextSnippet = BuildContextSnippet(text, start, end)
			};
			issue.Suggestions.Add(new SpellSuggestion(
				"“",
				1.0,
				"Đổi dấu nháy thẳng (\") sang dấu nháy “"));
			issue.Suggestions.Add(new SpellSuggestion(
				"”",
				0.9,
				"Đổi dấu nháy thẳng (\") sang dấu đóng nháy kép ”"));
			issues.Add(issue);
		}

		// 11. Dấu gạch ngang dài – (en-dash / em-dash) đổi sang - (Type = Variant -> "Dấu gạch (–)")
		foreach (Match m in EnDashPattern.Matches(text))
		{
			int start = m.Index;
			int end = start + m.Length;
			if (OverlapsAny(claimedSpans, start, end))
			{
				continue;
			}
			claimedSpans.Add(new SpanRange(start, end));

			SpellCheckIssue issue = new SpellCheckIssue
			{
				Start = start,
				End = end,
				ParagraphIndex = paragraphIndex,
				Original = m.Value,
				Type = SpellCheckIssueType.Variant,
				Severity = SpellCheckSeverity.Warning,
				RuleId = "SPACE_EN_DASH",
				Source = "Quy tắc trình bày văn bản hành chính & văn bản Đảng",
				Description = $"Phát hiện dấu gạch ngang dài '{m.Value}'. Đổi sang dấu gạch nối chuẩn '-'.",
				ContextSnippet = BuildContextSnippet(text, start, end)
			};
			issue.Suggestions.Add(new SpellSuggestion(
				"-",
				1.0,
				$"Đổi dấu '{m.Value}' thành dấu '-'"));
			issues.Add(issue);
		}

		issues.Sort((SpellCheckIssue a, SpellCheckIssue b) => a.Start.CompareTo(b.Start));
		return issues;
	}

	public static string FormatVisibleSpaces(string input)
	{
		if (input == null)
		{
			return string.Empty;
		}
		if (input.Length == 0)
		{
			return "(Xóa bỏ)";
		}
		return input
			.Replace(' ', '␣')
			.Replace('\u00A0', '␣')
			.Replace("\t", "⇥");
	}

	private static bool OverlapsAny(List<SpanRange> ranges, int start, int end)
	{
		for (int i = 0; i < ranges.Count; i++)
		{
			if (ranges[i].Overlaps(start, end))
			{
				return true;
			}
		}
		return false;
	}

	private static string BuildContextSnippet(string text, int start, int end)
	{
		int ctxStart = Math.Max(0, start - 28);
		int ctxEnd = Math.Min(text.Length, end + 28);
		string prefix = ((ctxStart > 0) ? "…" : string.Empty);
		string suffix = ((ctxEnd < text.Length) ? "…" : string.Empty);
		string before = text.Substring(ctxStart, start - ctxStart);
		string middle = FormatVisibleSpaces(text.Substring(start, end - start));
		string after = text.Substring(end, ctxEnd - end);
		return (prefix + before + "[" + middle + "]" + after + suffix)
			.Replace('\r', ' ')
			.Replace('\n', ' ')
			.Replace('\a', ' ');
	}
}
